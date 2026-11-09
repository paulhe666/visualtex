//! The screen copy of an XDV/DVI file: the (x)dvipdfmx specials of each page — PDF code (TikZ/pgf, tcolorbox,
//! colour, transformations, images, form XObjects, shadings, soft masks, opacity) — are carried out the way
//! dvipdfmx carries them out and written as `dvisvgm:raw` SVG, so that dvisvgm draws the same page the PDF of the
//! same run shows. Positions, line breaks and page breaks are TeX's (the DVI is the same); only the specials
//! change. What cannot be drawn this way is reported per page (the page is then shown from the PDF).
//!
//! Coordinates (all in bp). A DVI position (h, v) is A = (h, −v) for dvipdfmx ("user" space); a page starts
//! with CTM = Tr(72, H − 72) (the 1in offset, H the page height), a form XObject with the identity. dvipdfmx
//! draws text at CTM·(A − c) (c: the coordinate stack's top, see `pdf:bcontent`). dvisvgm puts the glyph at
//! S = (h, v) = K(A), K = (x, y) ↦ (x, −y). With F: device → SVG of the frame (a page: (x, y) ↦
//! (x − 72, H − 72 − y); a form: K), the SVG group around dvisvgm's text has the transform
//! G = F·CTM·Tr(−c)·K ("E group"), and anything drawn at CTM·X goes into that group at K(X + c).

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use crate::dvi::{self, Dvi, FontKind, Op, Reader};
use crate::pdfobj::{tokens, Lexer, Obj, Tok};
use crate::tfm::Metrics;
use crate::util::{num, xml_escape, Fnv, Mat};

#[derive(Clone, Copy, PartialEq, Debug)]
pub struct Rgb(pub f64, pub f64, pub f64);

impl Rgb {
    pub const BLACK: Rgb = Rgb(0.0, 0.0, 0.0);
    pub fn gray(g: f64) -> Rgb {
        Rgb(g, g, g)
    }
    pub fn cmyk(c: f64, m: f64, y: f64, k: f64) -> Rgb {
        // as MuPDF converts without a colour profile
        Rgb(1.0 - (c + k).min(1.0), 1.0 - (m + k).min(1.0), 1.0 - (y + k).min(1.0))
    }
    pub fn from_components(v: &[f64]) -> Rgb {
        match v.len() {
            1 => Rgb::gray(v[0]),
            3 => Rgb(v[0], v[1], v[2]),
            4 => Rgb::cmyk(v[0], v[1], v[2], v[3]),
            _ => Rgb::BLACK,
        }
    }
    pub fn hex(&self) -> String {
        let c = |v: f64| (v.clamp(0.0, 1.0) * 255.0).round() as u8;
        format!("#{:02x}{:02x}{:02x}", c(self.0), c(self.1), c(self.2))
    }
    fn is_black(&self) -> bool {
        self.hex() == "#000000"
    }
}

/// An image the host found for a `pdf:image` special.
pub enum ImageSource {
    /// a bitmap (PNG, JPEG): its pixel size and density (bp per pixel)
    Raster { width: f64, height: f64, xdensity: f64, ydensity: f64 },
    /// a PDF/EPS page drawn into `href`, covering this box of its own coordinates
    Box { llx: f64, lly: f64, urx: f64, ury: f64 },
}

pub struct ImageData {
    pub source: ImageSource,
    /// a data: URL
    pub href: String,
}

pub trait Host {
    /// The image file named in a special (relative to the build folder), page `page` of a PDF.
    fn image(&mut self, file: &str, page: u32) -> Option<Arc<ImageData>>;
}

pub struct ScreenPage {
    pub body: Vec<u8>,
    /// page size in bp
    pub width: f64,
    pub height: f64,
    /// what could not be drawn (empty: the page is exact)
    pub unsupported: Vec<String>,
    pub hash: u64,
}

#[derive(Clone, Copy, Default, Debug)]
struct Regs {
    h: f64,
    v: f64,
    w: f64,
    x: f64,
    y: f64,
    z: f64,
    f: u32,
}

#[derive(Clone, Debug)]
struct GState {
    ctm: Mat,
    fill: Rgb,
    stroke: Rgb,
    fill_alpha: f64,
    stroke_alpha: f64,
    width: f64,
    cap: u8,
    join: u8,
    miter: f64,
    dash: (Vec<f64>, f64),
    blend: Option<String>,
    fill_components: usize,
    stroke_components: usize,
    /// a tiling pattern painting fills / strokes (the SVG pattern's id)
    fill_paint: Option<String>,
    stroke_paint: Option<String>,
    /// clip/mask groups open for this state
    depth: usize,
}

impl GState {
    fn new(fill: Rgb, stroke: Rgb) -> GState {
        GState {
            ctm: Mat::ID,
            fill,
            stroke,
            fill_alpha: 1.0,
            stroke_alpha: 1.0,
            width: 1.0,
            cap: 0,
            join: 0,
            miter: 10.0,
            dash: (vec![], 0.0),
            blend: None,
            fill_components: 1,
            stroke_components: 1,
            fill_paint: None,
            stroke_paint: None,
            depth: 0,
        }
    }
}

/// dvipdfmx's dimension/transformation options (`width 3cm rotate 30 bbox 0 0 10 10 clip 1 …`).
#[derive(Clone, Debug)]
struct Ti {
    width: Option<f64>,
    height: Option<f64>,
    depth: f64,
    bbox: Option<[f64; 4]>,
    matrix: Mat,
    clip: bool,
    hide: bool,
    page: u32,
}

impl Ti {
    fn none() -> Ti {
        Ti { width: None, height: None, depth: 0.0, bbox: None, matrix: Mat::ID, clip: false, hide: false, page: 1 }
    }
    fn has_height(&self) -> bool {
        self.height.is_some()
    }
}

/// `spc_read_dimtrns_pdfm`: the options before an image's file name or after a form's name.
fn read_ti(lx: &mut Lexer) -> Ti {
    let mut t = Ti::none();
    t.clip = true;
    let (mut xs, mut ys, mut rot) = (1.0, 1.0, 0.0);
    let (mut has_scale, mut has_xs, mut has_ys, mut has_matrix) = (false, false, false, false);
    while let Some(key) = lx.ident() {
        match key.as_str() {
            "width" => t.width = lx.length(),
            "height" => t.height = Some(lx.length().unwrap_or(0.0)),
            "depth" => {
                t.depth = lx.length().unwrap_or(0.0);
                if t.height.is_none() {
                    t.height = Some(0.0);
                }
            }
            "scale" => {
                let v = lx.number().unwrap_or(1.0);
                xs = v;
                ys = v;
                has_scale = true;
            }
            "xscale" => {
                xs = lx.number().unwrap_or(1.0);
                has_xs = true;
            }
            "yscale" => {
                ys = lx.number().unwrap_or(1.0);
                has_ys = true;
            }
            "rotate" => rot = lx.number().unwrap_or(0.0).to_radians(),
            "bbox" => {
                let v: Vec<f64> = (0..4).filter_map(|_| lx.number()).collect();
                if v.len() == 4 {
                    t.bbox = Some([v[0], v[1], v[2], v[3]]);
                }
            }
            "matrix" => {
                let v: Vec<f64> = (0..6).filter_map(|_| lx.number()).collect();
                if v.len() == 6 {
                    t.matrix = Mat::new(v[0], v[1], v[2], v[3], v[4], v[5]);
                    has_matrix = true;
                }
            }
            "clip" => t.clip = lx.number().unwrap_or(1.0) != 0.0,
            "hide" => t.hide = true,
            "page" => t.page = lx.number().unwrap_or(1.0).max(1.0) as u32,
            "pagebox" => {
                lx.ident();
            }
            _ => break,
        }
    }
    if has_xs && t.width.is_some() {
        xs = 1.0;
    }
    if has_ys && t.has_height() {
        ys = 1.0;
    }
    let _ = has_scale;
    if !has_matrix {
        let (c, s) = (rot.cos(), rot.sin());
        t.matrix = Mat::new(xs * c, xs * s, -ys * s, ys * c, 0.0, 0.0);
    }
    if t.bbox.is_none() {
        t.clip = false;
    }
    t
}

/// A captured form XObject (`pdf:bxobj … pdf:exobj`), or a named image (`pdf:image @name …`).
struct Form {
    /// the page operations between bxobj and exobj
    bytes: Vec<u8>,
    regs: Regs,
    coord: (f64, f64),
    /// the form's reference point (A − c at bxobj), P space
    refp: (f64, f64),
    /// its box relative to the reference point
    bbox: [f64; 4],
    resources: Obj,
    fill: Rgb,
    stroke: Rgb,
    balanced: bool,
    image: Option<(Arc<ImageData>, Ti)>,
}

struct Capture {
    name: String,
    regs: Regs,
    coord: (f64, f64),
    refp: (f64, f64),
    bbox: [f64; 4],
    bytes: Vec<u8>,
    fill: Rgb,
    stroke: Rgb,
    /// the drawing state of the frame before bxobj (restored at exobj)
    saved: (GState, Vec<GState>, usize, Option<String>, Obj),
}

/// The drawing state of a page, or of a form being drawn.
struct Frame {
    gs: GState,
    stack: Vec<GState>,
    coords: Vec<(f64, f64)>,
    /// clip/mask groups open
    groups: usize,
    /// the open E group's attributes (None: no E group open)
    e_open: Option<String>,
    path: String,
    cur: (f64, f64),
    start: (f64, f64),
    pending_clip: Option<bool>,
    /// device → SVG of this frame
    fm: Mat,
    resources: Obj,
    regs: Regs,
    dstack: Vec<Regs>,
    captures: Vec<Capture>,
    depth: usize,
    /// inside a captured definition (its own drawing is hidden)
    hidden: usize,
    in_text: bool,
}

impl Frame {
    /// A page (`page_h`: its height) or, with None, a form.
    fn new(page_h: Option<f64>, fill: Rgb, stroke: Rgb, regs: Regs, coords: Vec<(f64, f64)>, resources: Obj, depth: usize) -> Frame {
        let mut gs = GState::new(fill, stroke);
        let fm = match page_h {
            Some(h) => {
                gs.ctm = Mat::translate(72.0, h - 72.0);
                Mat::new(1.0, 0.0, 0.0, -1.0, -72.0, h - 72.0)
            }
            None => K,
        };
        Frame {
            gs,
            stack: vec![],
            coords,
            groups: 0,
            e_open: None,
            path: String::new(),
            cur: (0.0, 0.0),
            start: (0.0, 0.0),
            pending_clip: None,
            fm,
            resources,
            regs,
            dstack: vec![],
            captures: vec![],
            depth,
            hidden: 0,
            in_text: false,
        }
    }
    fn c(&self) -> (f64, f64) {
        *self.coords.last().unwrap_or(&(0.0, 0.0))
    }
    /// G = F·CTM·Tr(−c)·K
    fn g(&self) -> Mat {
        let c = self.c();
        self.fm.mul(&self.gs.ctm).mul(&Mat::translate(-c.0, -c.1)).mul(&K)
    }
    /// user space (at the CTM) → E group: K·Tr(c)
    fn user_to_e(&self) -> Mat {
        let c = self.c();
        K.mul(&Mat::translate(c.0, c.1))
    }
}

pub struct Emu<'a> {
    dvi: &'a Dvi,
    metrics: &'a Metrics,
    host: &'a mut dyn Host,
    unit: f64,
    objects: HashMap<String, Obj>,
    /// `pdf:stream` data by name (pattern cells)
    streams: HashMap<String, Vec<u8>>,
    forms: HashMap<String, Form>,
    /// dvipdfmx's colour stack: (stroke, fill); the bottom is black
    colors: Vec<(Rgb, Rgb)>,
    paper: (f64, f64),
    ids: usize,
    page_no: usize,
    unsupported: Vec<String>,
    fonts_used: HashSet<u32>,
    out: Vec<u8>,
}

const MAX_DEPTH: usize = 8;

/// (x, y) ↦ (x, −y): dvipdfmx's user space ↔ dvisvgm's.
const K: Mat = Mat { a: 1.0, b: 0.0, c: 0.0, d: -1.0, e: 0.0, f: 0.0 };

impl<'a> Emu<'a> {
    pub fn new(dvi: &'a Dvi, metrics: &'a Metrics, host: &'a mut dyn Host, paper: (f64, f64)) -> Emu<'a> {
        Emu {
            dvi,
            metrics,
            host,
            unit: dvi.unit_bp(),
            objects: HashMap::new(),
            streams: HashMap::new(),
            forms: HashMap::new(),
            colors: vec![(Rgb::BLACK, Rgb::BLACK)],
            paper,
            ids: 0,
            page_no: 0,
            unsupported: vec![],
            fonts_used: HashSet::new(),
            out: vec![],
        }
    }

    /// All pages, in order (the state some specials leave — colours, named objects — carries over).
    pub fn rewrite_all(&mut self) -> Vec<ScreenPage> {
        let mut pages = Vec::with_capacity(self.dvi.pages.len());
        for i in 0..self.dvi.pages.len() {
            pages.push(self.rewrite_page(i));
        }
        pages
    }

    fn rewrite_page(&mut self, index: usize) -> ScreenPage {
        let dvi = self.dvi;
        let body = &dvi.pages[index].body;
        self.page_no = index + 1;
        self.ids = 0;
        self.unsupported.clear();
        self.fonts_used.clear();
        self.out = Vec::with_capacity(body.len() + 256);
        // the page size and background, from the specials anywhere on the page
        let mut background: Option<Rgb> = None;
        {
            let mut r = Reader::new(body);
            while r.pos < body.len() {
                match r.op() {
                    Ok(Op::Xxx(t)) => {
                        let s = String::from_utf8_lossy(t);
                        let s = s.trim_start();
                        if let Some(size) = page_size_special(s) {
                            self.paper = size;
                        } else if let Some(rest) = s.strip_prefix("background") {
                            background = parse_color_spec(rest.trim());
                        } else if let Some(rest) = strip_pdf(s).and_then(|r| r.strip_prefix("bgcolor").or_else(|| r.strip_prefix("bgc ")).or_else(|| r.strip_prefix("bbc "))) {
                            background = parse_pdf_colors(rest).map(|(f, _)| f);
                        }
                    }
                    Ok(_) => {}
                    Err(_) => break,
                }
            }
        }
        let (w, h) = self.paper;
        let top = *self.colors.last().unwrap();
        let mut fr = Frame::new(Some(h), top.1, top.0, Regs::default(), vec![], Obj::Dict(vec![]), 0);
        self.raw_special(&format!("dvisvgm:bbox fix -72bp -72bp {}bp {}bp", num(w - 72.0), num(h - 72.0)));
        if let Some(bg) = background {
            self.raw(&format!("<rect x='-72' y='-72' width='{}' height='{}' fill='{}'/>", num(w), num(h), bg.hex()));
        }
        self.sync_e(&mut fr);
        self.run(&mut fr, body);
        self.finish_frame(&mut fr);
        let body_out = std::mem::take(&mut self.out);
        let mut hash = Fnv::default();
        hash.write(&body_out);
        let mut fonts: Vec<&u32> = self.fonts_used.iter().collect();
        fonts.sort();
        for f in fonts {
            if let Some(def) = dvi.fonts.get(f) {
                hash.write(&def.raw);
            }
        }
        hash.write(format!("{w} {h}").as_bytes());
        ScreenPage { body: body_out, width: w, height: h, unsupported: std::mem::take(&mut self.unsupported), hash: hash.finish() }
    }

    fn flag(&mut self, what: &str) {
        if !self.unsupported.iter().any(|u| u == what) {
            self.unsupported.push(what.to_string());
        }
    }

    fn raw_special(&mut self, text: &str) {
        dvi::xxx(&mut self.out, text.as_bytes());
    }

    fn raw(&mut self, svg: &str) {
        let mut s = String::with_capacity(svg.len() + 12);
        s.push_str("dvisvgm:raw ");
        s.push_str(svg);
        dvi::xxx(&mut self.out, s.as_bytes());
    }

    fn id(&mut self) -> String {
        self.ids += 1;
        format!("vtx{}", self.ids)
    }

    // ---------------------------------------------------------------- groups --

    fn e_attrs(&self, fr: &Frame) -> String {
        let mut a = String::new();
        let g = fr.g();
        if !g.is_identity() {
            a.push_str(&format!(" transform='{}'", g.svg()));
        }
        if !fr.gs.fill.is_black() {
            a.push_str(&format!(" fill='{}'", fr.gs.fill.hex()));
        }
        if fr.gs.fill_alpha < 1.0 {
            a.push_str(&format!(" fill-opacity='{}'", num(fr.gs.fill_alpha)));
        }
        if fr.gs.stroke_alpha < 1.0 {
            a.push_str(&format!(" stroke-opacity='{}'", num(fr.gs.stroke_alpha)));
        }
        if let Some(b) = &fr.gs.blend {
            a.push_str(&format!(" style='mix-blend-mode:{b}'"));
        }
        a
    }

    /// The E group as the state says (closed and opened again when it changed).
    fn sync_e(&mut self, fr: &mut Frame) {
        let attrs = self.e_attrs(fr);
        if fr.e_open.as_deref() == Some(attrs.as_str()) {
            return;
        }
        self.close_e(fr);
        if !attrs.is_empty() {
            self.raw(&format!("<g{attrs}>"));
            fr.e_open = Some(attrs);
        }
    }

    fn close_e(&mut self, fr: &mut Frame) {
        if fr.e_open.take().is_some() {
            self.raw("</g>");
        }
    }

    /// A clip or mask group: `defs` (the clipPath/mask) and the opening tag.
    fn push_group(&mut self, fr: &mut Frame, defs: &str, open: &str) {
        self.close_e(fr);
        self.raw(&format!("{defs}{open}"));
        fr.groups += 1;
        fr.gs.depth = fr.groups;
    }

    fn close_groups_to(&mut self, fr: &mut Frame, depth: usize) {
        if fr.groups > depth {
            self.close_e(fr);
            while fr.groups > depth {
                self.raw("</g>");
                fr.groups -= 1;
            }
        }
    }

    fn finish_frame(&mut self, fr: &mut Frame) {
        while let Some(cap) = fr.captures.pop() {
            // a form never ended: what it drew stays hidden
            self.close_groups_to(fr, 0);
            self.raw("</g>");
            let (gs, stack, groups, e, res) = cap.saved;
            fr.gs = gs;
            fr.stack = stack;
            fr.groups = groups;
            fr.e_open = e;
            fr.resources = res;
            self.flag("unfinished form XObject");
        }
        self.close_groups_to(fr, 0);
        self.close_e(fr);
    }

    // ---------------------------------------------------------------- DVI --

    /// The DVI position in dvipdfmx's user space.
    fn pos_p(&self, fr: &Frame) -> (f64, f64) {
        (fr.regs.h * self.unit, -fr.regs.v * self.unit)
    }

    /// The page operations, with the specials carried out.
    fn run(&mut self, fr: &mut Frame, body: &[u8]) {
        let mut r = Reader::new(body);
        while r.pos < body.len() {
            let start = r.pos;
            let op = match r.op() {
                Ok(op) => op,
                Err(_) => {
                    self.flag("unreadable page operations");
                    break;
                }
            };
            let raw = &body[start..r.pos];
            if let Op::Xxx(text) = op {
                self.special(fr, text, raw);
                continue;
            }
            for cap in fr.captures.iter_mut() {
                cap.bytes.extend_from_slice(raw);
            }
            self.out.extend_from_slice(raw);
            let reg = &mut fr.regs;
            match op {
                Op::SetChar(c) => {
                    reg.h += self.char_width(reg.f, c);
                }
                Op::SetRule(_, b) => reg.h += b as f64,
                Op::Push => fr.dstack.push(*reg),
                Op::Pop => {
                    if let Some(s) = fr.dstack.pop() {
                        let f = reg.f;
                        *reg = s;
                        reg.f = f;
                    }
                }
                Op::Right(b) => reg.h += b as f64,
                Op::W0 => reg.h += reg.w,
                Op::W(b) => {
                    reg.w = b as f64;
                    reg.h += reg.w;
                }
                Op::X0 => reg.h += reg.x,
                Op::X(b) => {
                    reg.x = b as f64;
                    reg.h += reg.x;
                }
                Op::Down(a) => reg.v += a as f64,
                Op::Y0 => reg.v += reg.y,
                Op::Y(a) => {
                    reg.y = a as f64;
                    reg.v += reg.y;
                }
                Op::Z0 => reg.v += reg.z,
                Op::Z(a) => {
                    reg.z = a as f64;
                    reg.v += reg.z;
                }
                Op::Fnt(k) => {
                    reg.f = k;
                    self.fonts_used.insert(k);
                }
                Op::FntDef(k, _) => {
                    self.fonts_used.insert(k);
                }
                Op::Glyphs(w) => reg.h += w as f64,
                Op::Pic => self.flag("XeTeX picture"),
                Op::PutChar(_) | Op::PutRule(_, _) | Op::Nop | Op::Dir(_) | Op::Xxx(_) => {}
            }
        }
    }

    fn char_width(&self, font: u32, ch: u32) -> f64 {
        match self.dvi.fonts.get(&font).map(|d| &d.kind) {
            Some(FontKind::Tfm { .. }) => self.metrics.width(font, ch),
            _ => 0.0,
        }
    }

    // ---------------------------------------------------------------- specials --

    fn special(&mut self, fr: &mut Frame, text: &[u8], raw: &[u8]) {
        let s = String::from_utf8_lossy(text).to_string();
        let st = s.trim_start();
        let pdf = strip_pdf(st);
        let command = pdf.map(|p| p.split(|c: char| c.is_whitespace() || c == '<' || c == '[' || c == '@' || c == '(').next().unwrap_or(""));
        let is_exobj = matches!(command, Some("exobj" | "eform" | "endxobj"));
        if is_exobj {
            self.end_form(fr);
            for cap in fr.captures.iter_mut() {
                cap.bytes.extend_from_slice(raw);
            }
            return;
        }
        for cap in fr.captures.iter_mut() {
            cap.bytes.extend_from_slice(raw);
        }
        if let (Some(rest), Some(cmd)) = (pdf, command) {
            let args = rest[cmd.len()..].trim_start();
            self.pdf_special(fr, cmd, args);
        } else if let Some(rest) = st.strip_prefix("color") {
            self.color_special(fr, rest.trim());
        } else if let Some(rest) = st.strip_prefix("x:") {
            self.x_special(fr, rest.trim());
        } else if st.starts_with("dvisvgm:") || st.starts_with("em:") || st.starts_with("pn ") || st.starts_with("pa ") || st.starts_with("fp") || st.starts_with("ip") || st.starts_with("da ") || st.starts_with("dt ") || st.starts_with("sp") || st.starts_with("ar ") || st.starts_with("ia ") || st.starts_with("sh ") || st.starts_with("wh") || st.starts_with("bk") || st.starts_with("tx ") {
            // dvisvgm's own and the tpic/emTeX drawing specials both programs understand
            self.out.extend_from_slice(raw);
        } else if let Some(rest) = st.strip_prefix("PSfile").or_else(|| st.strip_prefix("psfile")) {
            self.psfile_special(fr, rest);
        } else if st.starts_with("ps:") || st.starts_with('"') || st.starts_with("!") {
            self.flag("PostScript special");
        }
        // everything else (papersize, background, hyperlinks, src:, html:, header=) is not drawn here
        self.sync_e(fr);
    }

    fn pdf_special(&mut self, fr: &mut Frame, cmd: &str, args: &str) {
        let (x, y) = self.pos_p(fr);
        match cmd {
            "code" => self.content(fr, args.as_bytes()),
            "literal" => {
                let mut a = args;
                let mut direct = false;
                loop {
                    if let Some(r) = a.strip_prefix("direct") {
                        direct = true;
                        a = r.trim_start();
                    } else if let Some(r) = a.strip_prefix("page") {
                        direct = true;
                        a = r.trim_start();
                    } else if let Some(r) = a.strip_prefix("reverse") {
                        a = r.trim_start();
                    } else {
                        break;
                    }
                }
                if direct {
                    self.content(fr, a.as_bytes());
                } else {
                    let c = fr.c();
                    let (cx, cy) = (x - c.0, y - c.1);
                    fr.gs.ctm = fr.gs.ctm.mul(&Mat::translate(cx, cy));
                    self.content(fr, a.as_bytes());
                    fr.gs.ctm = fr.gs.ctm.mul(&Mat::translate(-cx, -cy));
                }
            }
            "content" => {
                if !args.trim().is_empty() {
                    let c = fr.c();
                    self.op_q(fr);
                    fr.gs.ctm = fr.gs.ctm.mul(&Mat::translate(x - c.0, y - c.1));
                    self.content(fr, args.as_bytes());
                    self.op_big_q(fr);
                }
            }
            "bcontent" => {
                self.op_q(fr);
                let c = fr.c();
                fr.gs.ctm = fr.gs.ctm.mul(&Mat::translate(x - c.0, y - c.1));
                fr.coords.push((x, y));
            }
            "econtent" => {
                fr.coords.pop();
                self.op_big_q(fr);
                self.reset_color(fr);
            }
            "btrans" | "begintransform" | "begintrans" | "bt" => {
                let mut lx = Lexer::new(args.as_bytes());
                let ti = read_ti(&mut lx);
                let c = fr.c();
                let (px, py) = (x - c.0, y - c.1);
                let mut m = ti.matrix;
                m.e += (1.0 - m.a) * px - m.c * py;
                m.f += (1.0 - m.d) * py - m.b * px;
                self.op_q(fr);
                fr.gs.ctm = fr.gs.ctm.mul(&m);
            }
            "etrans" | "endtransform" | "endtrans" | "et" => {
                self.op_big_q(fr);
                self.reset_color(fr);
            }
            "bcolor" | "bc" | "begincolor" | "bgray" | "bg" | "begingray" => {
                let top = *self.colors.last().unwrap();
                if let Some((f, s)) = parse_pdf_colors_with(args, top) {
                    self.colors.push((s, f));
                    self.reset_color(fr);
                }
            }
            "scolor" | "sc" | "setcolor" => {
                let top = *self.colors.last().unwrap();
                if let Some((f, s)) = parse_pdf_colors_with(args, top) {
                    *self.colors.last_mut().unwrap() = (s, f);
                    self.reset_color(fr);
                }
            }
            "ecolor" | "ec" | "endcolor" | "egray" | "eg" | "endgray" => {
                if self.colors.len() > 1 {
                    self.colors.pop();
                }
                self.reset_color(fr);
            }
            "obj" | "object" => {
                let mut lx = Lexer::new(args.as_bytes());
                if let Some(Obj::Ref(name)) = lx.object() {
                    let value = lx.object().unwrap_or(Obj::Null);
                    self.objects.insert(name, value);
                }
            }
            "put" => {
                let mut lx = Lexer::new(args.as_bytes());
                if let Some(Obj::Ref(name)) = lx.object() {
                    let value = lx.object().unwrap_or(Obj::Null);
                    if name == "resources" {
                        fr.resources.merge(&value);
                    } else if let Some(o) = self.objects.get_mut(&name) {
                        o.merge(&value);
                    } else {
                        self.objects.insert(name, value);
                    }
                }
            }
            "pageresources" => {
                let mut lx = Lexer::new(args.as_bytes());
                if let Some(v) = lx.object() {
                    fr.resources.merge(&v);
                }
            }
            "stream" | "fstream" => {
                let mut lx = Lexer::new(args.as_bytes());
                if let Some(Obj::Ref(name)) = lx.object() {
                    let data = match lx.object() {
                        Some(Obj::Str(d)) => d,
                        _ => vec![],
                    };
                    let dict = lx.object().unwrap_or(Obj::Dict(vec![]));
                    self.streams.insert(name.clone(), data);
                    self.objects.insert(name, dict);
                }
            }
            "image" | "img" | "epdf" => self.image_special(fr, args),
            "bxobj" | "bform" | "beginxobj" => self.begin_form(fr, args),
            "uxobj" | "usexobj" => {
                let mut lx = Lexer::new(args.as_bytes());
                if let Some(Obj::Ref(name)) = lx.object() {
                    let ti = if lx.at_end() { Ti::none() } else { read_ti(&mut lx) };
                    let a = self.pos_p(fr);
                    self.use_form(fr, &name, &ti, a);
                } else {
                    self.flag("pdf:uxobj without a name");
                }
            }
            "mapline" | "mapfile" => {
                // font maps: dvisvgm reads these too
                let text = format!("pdf:{cmd} {args}");
                self.raw_special(&text);
            }
            "bxgstate" | "exgstate" | "bop" | "eop" => self.flag(&format!("pdf:{cmd}")),
            _ => {}
        }
    }

    fn reset_color(&mut self, fr: &mut Frame) {
        let (s, f) = *self.colors.last().unwrap();
        fr.gs.fill = f;
        fr.gs.stroke = s;
        fr.gs.fill_components = 3;
        fr.gs.stroke_components = 3;
    }

    fn color_special(&mut self, fr: &mut Frame, rest: &str) {
        if let Some(spec) = rest.strip_prefix("push") {
            let c = parse_color_spec(spec.trim()).unwrap_or(Rgb::BLACK);
            self.colors.push((c, c));
        } else if rest.starts_with("pop") {
            if self.colors.len() > 1 {
                self.colors.pop();
            }
        } else if let Some(c) = parse_color_spec(rest) {
            self.colors.truncate(1);
            self.colors[0] = (c, c);
        }
        self.reset_color(fr);
    }

    fn x_special(&mut self, fr: &mut Frame, rest: &str) {
        let (x, y) = self.pos_p(fr);
        let mut parts = rest.split_whitespace();
        let cmd = parts.next().unwrap_or("");
        let nums: Vec<f64> = parts.filter_map(|p| p.parse().ok()).collect();
        let around = |a: f64, b: f64, c: f64, d: f64| Mat::new(a, b, c, d, (1.0 - a) * x - c * y, (1.0 - d) * y - b * x);
        match cmd {
            "gsave" => self.op_q(fr),
            "grestore" => {
                self.op_big_q(fr);
                self.reset_color(fr);
            }
            "scale" | "bscale" if nums.len() >= 2 => fr.gs.ctm = fr.gs.ctm.mul(&around(nums[0], 0.0, 0.0, nums[1])),
            "escale" => {}
            "rotate" if !nums.is_empty() => {
                let r = nums[0].to_radians();
                fr.gs.ctm = fr.gs.ctm.mul(&around(r.cos(), r.sin(), -r.sin(), r.cos()));
            }
            "backgroundcolor" | "fontmapline" | "fontmapfile" | "papersize" => {}
            "" => {}
            other => self.flag(&format!("x:{other}")),
        }
    }

    // ---------------------------------------------------------------- PDF content --

    fn op_q(&mut self, fr: &mut Frame) {
        fr.stack.push(fr.gs.clone());
    }

    fn op_big_q(&mut self, fr: &mut Frame) {
        if let Some(saved) = fr.stack.pop() {
            let depth = saved.depth;
            fr.gs = saved;
            self.close_groups_to(fr, depth);
        }
    }

    /// PDF content operators (pdf:code and its kin).
    fn content(&mut self, fr: &mut Frame, code: &[u8]) {
        let toks = tokens(code);
        let mut operands: Vec<Obj> = Vec::new();
        let mut i = 0;
        while i < toks.len() {
            match &toks[i] {
                Tok::Operand(o) => operands.push(o.clone()),
                Tok::Op(op) => {
                    if fr.in_text {
                        if op == "ET" {
                            fr.in_text = false;
                        }
                    } else if op == "BI" {
                        self.flag("inline image in PDF code");
                        while i < toks.len() && toks[i] != Tok::Op("EI".into()) {
                            i += 1;
                        }
                    } else {
                        self.operator(fr, op, &operands);
                    }
                    operands.clear();
                }
            }
            i += 1;
        }
    }

    fn operator(&mut self, fr: &mut Frame, op: &str, a: &[Obj]) {
        let n = |i: usize| a.get(i).and_then(|o| o.num()).unwrap_or(0.0);
        let nums: Vec<f64> = a.iter().filter_map(|o| o.num()).collect();
        match op {
            "q" => self.op_q(fr),
            "Q" => self.op_big_q(fr),
            "cm" if nums.len() >= 6 => {
                fr.gs.ctm = fr.gs.ctm.mul(&Mat::new(nums[0], nums[1], nums[2], nums[3], nums[4], nums[5]));
            }
            "w" => fr.gs.width = n(0),
            "J" => fr.gs.cap = n(0) as u8,
            "j" => fr.gs.join = n(0) as u8,
            "M" => fr.gs.miter = n(0),
            "d" => {
                fr.gs.dash = (a.first().map(|o| o.nums()).unwrap_or_default(), n(1));
            }
            "ri" | "i" | "BMC" | "BDC" | "EMC" | "MP" | "DP" | "BX" | "EX" | "d0" | "d1" => {}
            "gs" => {
                if let Some(name) = a.first().and_then(|o| o.name()) {
                    self.ext_gstate(fr, name);
                }
            }
            "g" => {
                fr.gs.fill = Rgb::gray(n(0));
                fr.gs.fill_paint = None;
                fr.gs.fill_components = 1;
            }
            "G" => {
                fr.gs.stroke = Rgb::gray(n(0));
                fr.gs.stroke_paint = None;
                fr.gs.stroke_components = 1;
            }
            "rg" => {
                fr.gs.fill = Rgb(n(0), n(1), n(2));
                fr.gs.fill_paint = None;
                fr.gs.fill_components = 3;
            }
            "RG" => {
                fr.gs.stroke = Rgb(n(0), n(1), n(2));
                fr.gs.stroke_paint = None;
                fr.gs.stroke_components = 3;
            }
            "k" => {
                fr.gs.fill = Rgb::cmyk(n(0), n(1), n(2), n(3));
                fr.gs.fill_paint = None;
                fr.gs.fill_components = 4;
            }
            "K" => {
                fr.gs.stroke = Rgb::cmyk(n(0), n(1), n(2), n(3));
                fr.gs.stroke_paint = None;
                fr.gs.stroke_components = 4;
            }
            "cs" | "CS" => {
                let comps = match a.first().and_then(|o| o.name()) {
                    Some("DeviceGray") | Some("G") | Some("CalGray") => 1,
                    Some("DeviceRGB") | Some("RGB") | Some("CalRGB") => 3,
                    Some("DeviceCMYK") | Some("CMYK") => 4,
                    Some(other) => {
                        let cs = self.lookup_resource(fr, "ColorSpace", other);
                        match cs.as_ref().and_then(|o| o.array()).and_then(|v| v.first()).and_then(|o| o.name()) {
                            Some("Pattern") => 0,
                            Some("ICCBased") | Some("Separation") | Some("DeviceN") | Some("Indexed") | Some("Lab") => {
                                self.flag("colour space");
                                3
                            }
                            _ => match cs.as_ref().and_then(|o| o.name()) {
                                Some("DeviceGray") => 1,
                                Some("DeviceCMYK") => 4,
                                _ => 3,
                            },
                        }
                    }
                    None => 1,
                };
                if op == "cs" {
                    fr.gs.fill_components = comps;
                    fr.gs.fill = Rgb::BLACK;
                } else {
                    fr.gs.stroke_components = comps;
                    fr.gs.stroke = Rgb::BLACK;
                }
            }
            "sc" | "scn" | "SC" | "SCN" => {
                let fill = op.starts_with('s');
                if let Some(name) = a.iter().find_map(|o| o.name()) {
                    let color = Rgb::from_components(&nums);
                    match self.tiling_pattern(fr, name, color) {
                        Some(id) => {
                            if fill {
                                fr.gs.fill_paint = Some(id);
                            } else {
                                fr.gs.stroke_paint = Some(id);
                            }
                        }
                        None => self.flag("pattern"),
                    }
                    return;
                }
                let c = Rgb::from_components(&nums);
                if fill {
                    fr.gs.fill = c;
                    fr.gs.fill_paint = None;
                } else {
                    fr.gs.stroke = c;
                    fr.gs.stroke_paint = None;
                }
            }
            "m" if nums.len() >= 2 => {
                let p = self.pt(fr, nums[0], nums[1]);
                fr.path.push_str(&format!("M{} {}", num(p.0), num(p.1)));
                fr.cur = (nums[0], nums[1]);
                fr.start = fr.cur;
            }
            "l" if nums.len() >= 2 => {
                let p = self.pt(fr, nums[0], nums[1]);
                fr.path.push_str(&format!("L{} {}", num(p.0), num(p.1)));
                fr.cur = (nums[0], nums[1]);
            }
            "c" if nums.len() >= 6 => {
                self.curve(fr, (nums[0], nums[1]), (nums[2], nums[3]), (nums[4], nums[5]));
            }
            "v" if nums.len() >= 4 => {
                let c1 = fr.cur;
                self.curve(fr, c1, (nums[0], nums[1]), (nums[2], nums[3]));
            }
            "y" if nums.len() >= 4 => {
                self.curve(fr, (nums[0], nums[1]), (nums[2], nums[3]), (nums[2], nums[3]));
            }
            "h" => {
                fr.path.push('Z');
                fr.cur = fr.start;
            }
            "re" if nums.len() >= 4 => {
                let (x, y, w, h) = (nums[0], nums[1], nums[2], nums[3]);
                let pts = [(x, y), (x + w, y), (x + w, y + h), (x, y + h)];
                for (i, (px, py)) in pts.iter().enumerate() {
                    let p = self.pt(fr, *px, *py);
                    fr.path.push_str(&format!("{}{} {}", if i == 0 { 'M' } else { 'L' }, num(p.0), num(p.1)));
                }
                fr.path.push('Z');
                fr.cur = (x, y);
                fr.start = (x, y);
            }
            "W" => fr.pending_clip = Some(false),
            "W*" => fr.pending_clip = Some(true),
            "S" => self.paint(fr, false, true, false, false),
            "s" => self.paint(fr, false, true, false, true),
            "f" | "F" => self.paint(fr, true, false, false, false),
            "f*" => self.paint(fr, true, false, true, false),
            "B" => self.paint(fr, true, true, false, false),
            "B*" => self.paint(fr, true, true, true, false),
            "b" => self.paint(fr, true, true, false, true),
            "b*" => self.paint(fr, true, true, true, true),
            "n" => self.paint(fr, false, false, false, false),
            "sh" => {
                if let Some(name) = a.first().and_then(|o| o.name()) {
                    self.shade(fr, name);
                }
            }
            "Do" => {
                if let Some(name) = a.first().and_then(|o| o.name()) {
                    self.do_xobject(fr, name);
                }
            }
            "BT" => {
                fr.in_text = true;
                self.flag("text in PDF code");
            }
            "Tr" | "Tc" | "Tw" | "Tz" | "TL" | "Tf" | "Ts" => {
                if op == "Tr" && n(0) != 0.0 {
                    self.flag("text rendering mode");
                }
            }
            _ => {}
        }
    }

    /// A user-space point in the E group: F(u + c).
    fn pt(&self, fr: &Frame, x: f64, y: f64) -> (f64, f64) {
        fr.user_to_e().apply(x, y)
    }

    fn curve(&mut self, fr: &mut Frame, a: (f64, f64), b: (f64, f64), c: (f64, f64)) {
        let (p1, p2, p3) = (self.pt(fr, a.0, a.1), self.pt(fr, b.0, b.1), self.pt(fr, c.0, c.1));
        fr.path.push_str(&format!(
            "C{} {} {} {} {} {}",
            num(p1.0),
            num(p1.1),
            num(p2.0),
            num(p2.1),
            num(p3.0),
            num(p3.1)
        ));
        fr.cur = c;
    }

    fn stroke_attrs(&self, fr: &Frame) -> String {
        let gs = &fr.gs;
        let mut s = format!(" stroke='{}'", gs.stroke.hex());
        if gs.width <= 0.0 {
            s.push_str(" stroke-width='1' vector-effect='non-scaling-stroke'");
        } else if (gs.width - 1.0).abs() > 1e-9 {
            s.push_str(&format!(" stroke-width='{}'", num(gs.width)));
        }
        match gs.cap {
            1 => s.push_str(" stroke-linecap='round'"),
            2 => s.push_str(" stroke-linecap='square'"),
            _ => {}
        }
        match gs.join {
            1 => s.push_str(" stroke-linejoin='round'"),
            2 => s.push_str(" stroke-linejoin='bevel'"),
            _ => {}
        }
        if gs.join == 0 && (gs.miter - 4.0).abs() > 1e-9 {
            s.push_str(&format!(" stroke-miterlimit='{}'", num(gs.miter.max(1.0))));
        }
        if !gs.dash.0.is_empty() && gs.dash.0.iter().any(|v| *v > 0.0) {
            let list: Vec<String> = gs.dash.0.iter().map(|v| num(*v)).collect();
            s.push_str(&format!(" stroke-dasharray='{}'", list.join(" ")));
            if gs.dash.1 != 0.0 {
                s.push_str(&format!(" stroke-dashoffset='{}'", num(gs.dash.1)));
            }
        }
        s
    }

    fn paint(&mut self, fr: &mut Frame, fill: bool, stroke: bool, evenodd: bool, close: bool) {
        if close {
            fr.path.push('Z');
        }
        let d = std::mem::take(&mut fr.path);
        if d.is_empty() {
            fr.pending_clip = None;
            return;
        }
        self.sync_e(fr);
        if fill || stroke {
            let mut attrs = String::new();
            if fill {
                match &fr.gs.fill_paint {
                    Some(id) => attrs.push_str(&format!(" fill='url(#{id})'")),
                    None => attrs.push_str(&format!(" fill='{}'", fr.gs.fill.hex())),
                }
                if evenodd {
                    attrs.push_str(" fill-rule='evenodd'");
                }
            } else {
                attrs.push_str(" fill='none'");
            }
            if stroke {
                attrs.push_str(&self.stroke_attrs(fr));
                if let Some(id) = &fr.gs.stroke_paint {
                    attrs = attrs.replacen(&format!(" stroke='{}'", fr.gs.stroke.hex()), &format!(" stroke='url(#{id})'"), 1);
                }
            }
            self.raw(&format!("<path d='{d}'{attrs}/>"));
        }
        if let Some(eo) = fr.pending_clip.take() {
            let id = self.id();
            let g = fr.g();
            let tr = if g.is_identity() { String::new() } else { format!(" transform='{}'", g.svg()) };
            let rule = if eo { " clip-rule='evenodd'" } else { "" };
            let defs = format!("<clipPath id='{id}'><path d='{d}'{tr}{rule}/></clipPath>");
            self.push_group(fr, &defs, &format!("<g clip-path='url(#{id})'>"));
        }
    }

    fn lookup(&self, o: &Obj, depth: usize) -> Option<Obj> {
        match o {
            Obj::Ref(name) if depth < 8 => self.objects.get(name).and_then(|v| self.lookup(v, depth + 1)),
            Obj::Ref(_) => None,
            other => Some(other.clone()),
        }
    }

    /// A resource by category and name: the frame's resources, then any named object holding it.
    fn lookup_resource(&self, fr: &Frame, category: &str, name: &str) -> Option<Obj> {
        if let Some(cat) = fr.resources.get(category).and_then(|c| self.lookup(c, 0)) {
            if let Some(v) = cat.get(name) {
                return self.lookup(v, 0);
            }
        }
        // pgf puts the resources once (@pgfextgs) and names them in each page's /Resources
        for v in self.objects.values() {
            if let Some(found) = v.get(name) {
                return self.lookup(found, 0);
            }
        }
        None
    }

    fn ext_gstate(&mut self, fr: &mut Frame, name: &str) {
        let Some(d) = self.lookup_resource(fr, "ExtGState", name) else {
            self.flag("unknown graphics state");
            return;
        };
        if let Some(v) = d.get("CA").and_then(|o| o.num()) {
            fr.gs.stroke_alpha = v;
        }
        if let Some(v) = d.get("ca").and_then(|o| o.num()) {
            fr.gs.fill_alpha = v;
        }
        if let Some(v) = d.get("LW").and_then(|o| o.num()) {
            fr.gs.width = v;
        }
        if let Some(v) = d.get("LC").and_then(|o| o.num()) {
            fr.gs.cap = v as u8;
        }
        if let Some(v) = d.get("LJ").and_then(|o| o.num()) {
            fr.gs.join = v as u8;
        }
        if let Some(v) = d.get("ML").and_then(|o| o.num()) {
            fr.gs.miter = v;
        }
        if let Some(bm) = d.get("BM") {
            let name = bm.name().map(|s| s.to_string()).or_else(|| bm.array().and_then(|a| a.first()).and_then(|o| o.name()).map(|s| s.to_string()));
            fr.gs.blend = match name.as_deref() {
                None | Some("Normal") | Some("Compatible") => None,
                Some(other) => Some(blend_css(other)),
            };
        }
        if let Some(sm) = d.get("SMask") {
            let sm = self.lookup(sm, 0).unwrap_or(Obj::Null);
            if let Obj::Dict(_) = sm {
                match sm.get("G") {
                    Some(Obj::Ref(form)) => {
                        let luminosity = sm.get("S").and_then(|o| o.name()) != Some("Alpha");
                        if !luminosity {
                            self.flag("alpha soft mask");
                        }
                        self.soft_mask(fr, &form.clone());
                    }
                    _ => self.flag("soft mask"),
                }
            }
        }
    }

    // ---------------------------------------------------------------- patterns --

    /// A tiling pattern (PatternType 1, its cell a `pdf:stream`) as an SVG pattern in the E group's space;
    /// `color` paints an uncoloured one (PaintType 2).
    fn tiling_pattern(&mut self, fr: &mut Frame, name: &str, color: Rgb) -> Option<String> {
        let target = match fr.resources.get("Pattern").and_then(|c| self.lookup(c, 0)).and_then(|c| c.get(name).cloned()) {
            Some(Obj::Ref(r)) => r,
            _ => self.objects.values().find_map(|v| match v.get(name) {
                Some(Obj::Ref(r)) => Some(r.clone()),
                _ => None,
            })?,
        };
        let dict = self.objects.get(&target)?.clone();
        if dict.get("PatternType").and_then(|o| o.num()) != Some(1.0) {
            return None;
        }
        let data = self.streams.get(&target)?.clone();
        let bbox = self.lookup(dict.get("BBox")?, 0)?.nums();
        if bbox.len() < 4 {
            return None;
        }
        let xs = dict.get("XStep").and_then(|o| o.num()).unwrap_or(bbox[2] - bbox[0]);
        let ys = dict.get("YStep").and_then(|o| o.num()).unwrap_or(bbox[3] - bbox[1]);
        let m = dict.get("Matrix").and_then(|o| self.lookup(o, 0)).map(|o| o.nums()).filter(|v| v.len() == 6).map(|v| Mat::new(v[0], v[1], v[2], v[3], v[4], v[5])).unwrap_or(Mat::ID);
        let uncolored = dict.get("PaintType").and_then(|o| o.num()) == Some(2.0);
        // the cell drawn into a buffer of its own (the cell's PDF code as for a form at the origin)
        let saved = std::mem::take(&mut self.out);
        let mut cell = Frame::new(None, color, color, Regs::default(), vec![(0.0, 0.0)], Obj::Dict(vec![]), fr.depth + 1);
        self.content(&mut cell, &data);
        self.finish_frame(&mut cell);
        let raw = std::mem::replace(&mut self.out, saved);
        let mut svg = String::new();
        let mut r = Reader::new(&raw);
        while r.pos < raw.len() {
            if let Ok(Op::Xxx(t)) = r.op() {
                if let Some(x) = t.strip_prefix(b"dvisvgm:raw ") {
                    svg.push_str(&String::from_utf8_lossy(x));
                }
            } else {
                break;
            }
        }
        if uncolored {
            // colours inside an uncoloured cell do not count
            let hexes = ["#000000"];
            for h in hexes {
                svg = svg.replace(&format!("'{h}'"), &format!("'{}'", color.hex()));
            }
        }
        let id = self.id();
        let clip = self.id();
        // pattern space → E group: G⁻¹·K·M·K (the cell's SVG space is K of pattern space)
        let t = fr.g().invert().unwrap_or(Mat::ID).mul(&K).mul(&m).mul(&K);
        // the tile: one step, the cell drawn at the neighbouring steps too (a cell may reach past its step), each
        // clipped to the cell's box
        let (bx0, by0, bx1, by1) = (bbox[0], -bbox[3], bbox[2], -bbox[1]);
        let mut body = format!("<clipPath id='{clip}'><rect x='{}' y='{}' width='{}' height='{}'/></clipPath>", num(bx0), num(by0), num(bx1 - bx0), num(by1 - by0));
        for dy in [-1.0, 0.0, 1.0] {
            for dx in [-1.0, 0.0, 1.0] {
                body.push_str(&format!("<g transform='translate({} {})'><g clip-path='url(#{clip})'>{svg}</g></g>", num(dx * xs), num(-dy * ys)));
            }
        }
        self.sync_e(fr);
        self.raw(&format!(
            "<pattern id='{id}' patternUnits='userSpaceOnUse' x='0' y='{}' width='{}' height='{}' patternTransform='{}'>{body}</pattern>",
            num(-ys),
            num(xs),
            num(ys),
            t.svg()
        ));
        Some(id)
    }

    // ---------------------------------------------------------------- shadings --

    fn shade(&mut self, fr: &mut Frame, name: &str) {
        let Some(sh) = self.lookup_resource(fr, "Shading", name) else {
            self.flag("unknown shading");
            return;
        };
        let Some(id) = self.gradient(fr, &sh) else {
            self.flag("shading type");
            return;
        };
        self.sync_e(fr);
        self.raw(&format!("<rect x='-30000' y='-30000' width='60000' height='60000' fill='url(#{id})'/>"));
    }

    /// An axial or radial shading as an SVG gradient (defined here, used by its id).
    fn gradient(&mut self, fr: &mut Frame, sh: &Obj) -> Option<String> {
        let kind = sh.get("ShadingType")?.num()? as i32;
        if kind != 2 && kind != 3 {
            return None;
        }
        let coords = self.lookup(sh.get("Coords")?, 0)?.nums();
        let domain = sh.get("Domain").and_then(|d| self.lookup(d, 0)).map(|d| d.nums()).filter(|d| d.len() == 2).unwrap_or(vec![0.0, 1.0]);
        let function = self.lookup(sh.get("Function")?, 0)?;
        let space = sh.get("ColorSpace").and_then(|c| self.lookup(c, 0)).and_then(|c| c.name().map(|s| s.to_string())).unwrap_or("DeviceRGB".into());
        let stops = self.stops(&function, domain[0], domain[1])?;
        let id = self.id();
        let tr = fr.user_to_e().svg();
        let mut s = if kind == 2 && coords.len() >= 4 {
            format!(
                "<linearGradient id='{id}' gradientUnits='userSpaceOnUse' x1='{}' y1='{}' x2='{}' y2='{}' gradientTransform='{tr}'>",
                num(coords[0]),
                num(coords[1]),
                num(coords[2]),
                num(coords[3])
            )
        } else if coords.len() >= 6 {
            format!(
                "<radialGradient id='{id}' gradientUnits='userSpaceOnUse' fx='{}' fy='{}' fr='{}' cx='{}' cy='{}' r='{}' gradientTransform='{tr}'>",
                num(coords[0]),
                num(coords[1]),
                num(coords[2]),
                num(coords[3]),
                num(coords[4]),
                num(coords[5])
            )
        } else {
            return None;
        };
        for (offset, comps) in stops {
            let c = match space.as_str() {
                "DeviceGray" => Rgb::gray(*comps.first().unwrap_or(&0.0)),
                "DeviceCMYK" if comps.len() >= 4 => Rgb::cmyk(comps[0], comps[1], comps[2], comps[3]),
                _ => Rgb::from_components(&comps),
            };
            s.push_str(&format!("<stop offset='{}' stop-color='{}'/>", num(offset.clamp(0.0, 1.0)), c.hex()));
        }
        s.push_str(if kind == 2 { "</linearGradient>" } else { "</radialGradient>" });
        self.sync_e(fr);
        self.raw(&s);
        Some(id)
    }

    /// Gradient stops for a function over [t0, t1]: exact at the breakpoints of linear pieces, sampled otherwise.
    fn stops(&self, f: &Obj, t0: f64, t1: f64) -> Option<Vec<(f64, Vec<f64>)>> {
        let span = if (t1 - t0).abs() < 1e-12 { 1.0 } else { t1 - t0 };
        let mut ts: Vec<f64> = vec![t0, t1];
        let mut linear = true;
        self.breakpoints(f, t0, t1, &mut ts, &mut linear, 0);
        if !linear {
            ts = (0..=32).map(|i| t0 + (t1 - t0) * i as f64 / 32.0).collect();
        }
        ts.sort_by(|a, b| a.partial_cmp(b).unwrap());
        ts.dedup_by(|a, b| (*a - *b).abs() < 1e-9);
        let mut out = vec![];
        for (i, t) in ts.iter().enumerate() {
            // at a stitching breakpoint, both sides (a jump stays a jump)
            if i > 0 && i + 1 < ts.len() {
                let before = self.eval(f, t - 1e-7 * span.abs(), 0)?;
                let after = self.eval(f, *t, 0)?;
                out.push(((t - t0) / span, before));
                out.push(((t - t0) / span, after));
            } else {
                out.push(((t - t0) / span, self.eval(f, *t, 0)?));
            }
        }
        Some(out)
    }

    fn breakpoints(&self, f: &Obj, t0: f64, t1: f64, ts: &mut Vec<f64>, linear: &mut bool, depth: usize) {
        let Some(f) = self.lookup(f, 0) else { return };
        if depth > 6 {
            return;
        }
        match f.get("FunctionType").and_then(|o| o.num()).map(|v| v as i32) {
            Some(2) => {
                if (f.get("N").and_then(|o| o.num()).unwrap_or(1.0) - 1.0).abs() > 1e-9 {
                    *linear = false;
                }
            }
            Some(3) => {
                let bounds = f.get("Bounds").and_then(|o| self.lookup(o, 0)).map(|o| o.nums()).unwrap_or_default();
                let domain = f.get("Domain").and_then(|o| self.lookup(o, 0)).map(|o| o.nums()).unwrap_or(vec![t0, t1]);
                let encode = f.get("Encode").and_then(|o| self.lookup(o, 0)).map(|o| o.nums()).unwrap_or_default();
                for b in &bounds {
                    ts.push(*b);
                }
                if let Some(Obj::Array(fs)) = f.get("Functions").and_then(|o| self.lookup(o, 0)) {
                    let mut edges = vec![domain.first().copied().unwrap_or(t0)];
                    edges.extend(bounds.iter());
                    edges.push(domain.get(1).copied().unwrap_or(t1));
                    for (i, sub) in fs.iter().enumerate() {
                        let mut sub_ts = vec![];
                        self.breakpoints(sub, 0.0, 1.0, &mut sub_ts, linear, depth + 1);
                        if !sub_ts.is_empty() {
                            // nested stitching: its breakpoints mapped back
                            let (e0, e1) = (encode.get(2 * i).copied().unwrap_or(0.0), encode.get(2 * i + 1).copied().unwrap_or(1.0));
                            let (a, b) = (edges[i], edges.get(i + 1).copied().unwrap_or(t1));
                            for st in sub_ts {
                                if (e1 - e0).abs() > 1e-12 {
                                    ts.push(a + (st - e0) / (e1 - e0) * (b - a));
                                }
                            }
                        }
                    }
                }
            }
            _ => *linear = false,
        }
    }

    fn eval(&self, f: &Obj, t: f64, depth: usize) -> Option<Vec<f64>> {
        let f = self.lookup(f, 0)?;
        if depth > 8 {
            return None;
        }
        let domain = f.get("Domain").and_then(|o| self.lookup(o, 0)).map(|o| o.nums()).filter(|d| d.len() >= 2).unwrap_or(vec![0.0, 1.0]);
        let t = t.clamp(domain[0].min(domain[1]), domain[0].max(domain[1]));
        match f.get("FunctionType")?.num()? as i32 {
            2 => {
                let c0 = f.get("C0").and_then(|o| self.lookup(o, 0)).map(|o| o.nums()).unwrap_or(vec![0.0]);
                let c1 = f.get("C1").and_then(|o| self.lookup(o, 0)).map(|o| o.nums()).unwrap_or(vec![1.0]);
                let n = f.get("N").and_then(|o| o.num()).unwrap_or(1.0);
                let x = t.powf(n);
                Some(c0.iter().zip(c1.iter()).map(|(a, b)| a + x * (b - a)).collect())
            }
            3 => {
                let fs = f.get("Functions").and_then(|o| self.lookup(o, 0))?;
                let fs = fs.array()?;
                let bounds = f.get("Bounds").and_then(|o| self.lookup(o, 0)).map(|o| o.nums()).unwrap_or_default();
                let encode = f.get("Encode").and_then(|o| self.lookup(o, 0)).map(|o| o.nums()).unwrap_or_default();
                let mut k = 0;
                while k < bounds.len() && t >= bounds[k] {
                    k += 1;
                }
                let k = k.min(fs.len().saturating_sub(1));
                let lo = if k == 0 { domain[0] } else { bounds[k - 1] };
                let hi = if k < bounds.len() { bounds[k] } else { domain[1] };
                let (e0, e1) = (encode.get(2 * k).copied().unwrap_or(0.0), encode.get(2 * k + 1).copied().unwrap_or(1.0));
                let x = if (hi - lo).abs() < 1e-12 { e0 } else { e0 + (t - lo) / (hi - lo) * (e1 - e0) };
                self.eval(&fs[k], x, depth + 1)
            }
            _ => None,
        }
    }

    // ---------------------------------------------------------------- images and forms --

    fn image_special(&mut self, fr: &mut Frame, args: &str) {
        let mut lx = Lexer::new(args.as_bytes());
        lx.skip_white();
        let mut name = None;
        if lx.rest().first() == Some(&b'@') {
            if let Some(Obj::Ref(n)) = lx.object() {
                name = Some(n);
            }
        }
        let ti = read_ti(&mut lx);
        let file = match lx.object() {
            Some(Obj::Str(s)) => String::from_utf8_lossy(&s).to_string(),
            _ => {
                self.flag("image without a file name");
                return;
            }
        };
        let Some(img) = self.host.image(&file, ti.page) else {
            self.flag(&format!("image {file}"));
            return;
        };
        let a = self.pos_p(fr);
        if let Some(name) = name {
            self.forms.insert(
                name,
                Form {
                    bytes: vec![],
                    regs: fr.regs,
                    coord: fr.c(),
                    refp: (0.0, 0.0),
                    bbox: [0.0; 4],
                    resources: Obj::Null,
                    fill: fr.gs.fill,
                    stroke: fr.gs.stroke,
                    balanced: true,
                    image: Some((img.clone(), ti.clone())),
                },
            );
        }
        if !ti.hide {
            self.draw_image(fr, &img, &ti, a);
        }
    }

    /// `PSfile="x.eps" llx=… lly=… urx=… ury=… rwi=…` (graphics' dvips-style EPS inclusion, which xetex.def
    /// uses for EPS): dvipdfmx converts the file to a PDF form of its bounding box and places it.
    fn psfile_special(&mut self, fr: &mut Frame, rest: &str) {
        let rest = rest.trim_start().trim_start_matches('=').trim_start();
        let (file, args) = if let Some(r) = rest.strip_prefix('"') {
            match r.find('"') {
                Some(e) => (&r[..e], &r[e + 1..]),
                None => (r, ""),
            }
        } else {
            match rest.find(char::is_whitespace) {
                Some(e) => (&rest[..e], &rest[e..]),
                None => (rest, ""),
            }
        };
        let mut t = Ti::none();
        let (mut xo, mut yo, mut xs, mut ys, mut rot) = (0.0, 0.0, 1.0, 1.0, 0.0);
        let mut bbox = [0.0f64; 4];
        let mut has_bbox = false;
        for part in args.split_whitespace() {
            let (k, v) = match part.split_once('=') {
                Some((k, v)) => (k, v.trim_matches(|c| c == '"' || c == '\'').parse::<f64>().unwrap_or(0.0)),
                None => (part, 0.0),
            };
            match k {
                "hoffset" => xo = v,
                "voffset" => yo = v,
                "hsize" => t.width = Some(v),
                "vsize" => t.height = Some(v),
                "hscale" => xs = v / 100.0,
                "vscale" => ys = v / 100.0,
                "angle" => rot = v.to_radians(),
                "llx" | "lly" | "urx" | "ury" => {
                    let i = ["llx", "lly", "urx", "ury"].iter().position(|x| *x == k).unwrap();
                    bbox[i] = v;
                    has_bbox = true;
                }
                "rwi" => t.width = Some(v / 10.0),
                "rhi" => t.height = Some(v / 10.0),
                "clip" => t.clip = true,
                _ => {}
            }
        }
        if has_bbox {
            t.bbox = Some(bbox);
        } else {
            t.clip = false;
        }
        let (c, sn) = (rot.cos(), rot.sin());
        t.matrix = Mat::new(xs * c, xs * sn, -ys * sn, ys * c, xo, yo);
        let Some(img) = self.host.image(file, 1) else {
            self.flag(&format!("image {file}"));
            return;
        };
        let a = self.pos_p(fr);
        self.draw_image(fr, &img, &t, a);
    }

    /// An image placed at A (P space) with options `ti`: pdf_dev_put_image.
    fn draw_image(&mut self, fr: &mut Frame, img: &ImageData, ti: &Ti, a: (f64, f64)) {
        // T: the image's own space (unit square for bitmaps, its box for PDF/EPS) → placement; r: the clip
        let (fit, content, clip) = match &img.source {
            ImageSource::Raster { width, height, xdensity, ydensity } => {
                let (w, h) = (*width, *height);
                let (mut xd, mut yd) = (*xdensity, *ydensity);
                if let Some(b) = ti.bbox {
                    // the bbox graphicx gives (from extractbb) spans the whole image: its density
                    let (bw, bh) = (b[2] - b[0], b[3] - b[1]);
                    if b[0].abs() < 1e-6 && b[1].abs() < 1e-6 && bw > 0.0 && bh > 0.0 && ((bw / bh) - (w / h)).abs() < 0.02 * (w / h) {
                        xd = bw / w;
                        yd = bh / h;
                    }
                }
                let (wd0, ht0, xs, ys, dx, dy) = match ti.bbox {
                    Some(b) => {
                        let (wd0, ht0) = ((b[2] - b[0]).max(1e-9), (b[3] - b[1]).max(1e-9));
                        (wd0, ht0, w * xd / wd0, h * yd / ht0, -b[0] / wd0, -b[1] / ht0)
                    }
                    None => (w * xd, h * yd, 1.0, 1.0, 0.0, 0.0),
                };
                let (sx, sy, dp) = match (ti.width, ti.has_height()) {
                    (Some(wd), true) => (wd * xs, (ti.height.unwrap() + ti.depth) * ys, ti.depth * ys),
                    (Some(wd), false) => (wd * xs, wd * xs * (h / w), 0.0),
                    (None, true) => {
                        let sy = (ti.height.unwrap() + ti.depth) * ys;
                        (sy * (w / h), sy, ti.depth * ys)
                    }
                    (None, false) => (wd0, ht0, 0.0),
                };
                let fit = Mat::new(sx, 0.0, 0.0, sy, dx * sx / xs, dy * sy / ys - dp);
                let clip = if ti.clip {
                    ti.bbox.map(|b| [b[0] / (w * xd), b[1] / (h * yd), b[2] / (w * xd), b[3] / (h * yd)])
                } else {
                    None
                };
                (fit, Mat::ID, clip)
            }
            ImageSource::Box { llx, lly, urx, ury } => {
                let b = ti.bbox.unwrap_or([*llx, *lly, *urx, *ury]);
                let fit = fit_form(ti, b);
                (fit, Mat::new(urx - llx, 0.0, 0.0, ury - lly, *llx, *lly), if ti.clip { Some(b) } else { None })
            }
        };
        // E group: K·Tr(A)·ti·fit
        let place = K.mul(&Mat::translate(a.0, a.1)).mul(&ti.matrix).mul(&fit);
        let flip = Mat::new(1.0, 0.0, 0.0, -1.0, 0.0, 1.0);
        let t = place.mul(&content).mul(&flip);
        self.sync_e(fr);
        let mut close = false;
        if let Some(r) = clip {
            let id = self.id();
            let pts = [(r[0], r[1]), (r[2], r[1]), (r[2], r[3]), (r[0], r[3])];
            let d: Vec<String> = pts
                .iter()
                .enumerate()
                .map(|(i, (x, y))| {
                    let p = place.apply(*x, *y);
                    format!("{}{} {}", if i == 0 { 'M' } else { 'L' }, num(p.0), num(p.1))
                })
                .collect();
            self.raw(&format!("<clipPath id='{id}'><path d='{}Z'/></clipPath><g clip-path='url(#{id})'>", d.join("")));
            close = true;
        }
        let opacity = if fr.gs.fill_alpha < 1.0 { format!(" opacity='{}'", num(fr.gs.fill_alpha)) } else { String::new() };
        self.raw(&format!(
            "<image x='0' y='0' width='1' height='1' preserveAspectRatio='none' transform='{}'{opacity} xlink:href='{}'/>",
            t.svg(),
            xml_escape(&img.href)
        ));
        if close {
            self.raw("</g>");
        }
    }

    fn begin_form(&mut self, fr: &mut Frame, args: &str) {
        let mut lx = Lexer::new(args.as_bytes());
        let Some(Obj::Ref(name)) = lx.object() else {
            self.flag("pdf:bxobj without a name");
            return;
        };
        let ti = read_ti(&mut lx);
        let bbox = match ti.bbox {
            Some(b) => b,
            None => [0.0, -ti.depth, ti.width.unwrap_or(0.0), ti.height.unwrap_or(0.0)],
        };
        let a = self.pos_p(fr);
        let c = fr.c();
        let saved = (fr.gs.clone(), std::mem::take(&mut fr.stack), fr.groups, None, std::mem::replace(&mut fr.resources, Obj::Dict(vec![])));
        // what the definition draws is hidden where it stands (it is drawn where the form is used)
        self.close_e(fr);
        self.raw("<g display='none'>");
        fr.hidden += 1;
        let top = *self.colors.last().unwrap();
        let groups_before = fr.groups;
        fr.captures.push(Capture {
            name,
            regs: fr.regs,
            coord: c,
            refp: (a.0 - c.0, a.1 - c.1),
            bbox,
            bytes: vec![],
            fill: top.1,
            stroke: top.0,
            saved: (saved.0, saved.1, groups_before, saved.3, saved.4),
        });
        fr.gs = GState::new(top.1, top.0);
        fr.gs.depth = fr.groups;
        self.sync_e(fr);
    }

    fn end_form(&mut self, fr: &mut Frame) {
        let Some(cap) = fr.captures.pop() else {
            return;
        };
        let (gs, stack, groups, _e, res) = cap.saved;
        self.close_groups_to(fr, groups);
        self.close_e(fr);
        self.raw("</g>");
        fr.hidden = fr.hidden.saturating_sub(1);
        let form_resources = std::mem::replace(&mut fr.resources, res);
        fr.gs = gs;
        fr.stack = stack;
        fr.groups = groups;
        let balanced = balanced(&cap.bytes);
        self.forms.insert(
            cap.name,
            Form {
                bytes: cap.bytes,
                regs: cap.regs,
                coord: cap.coord,
                refp: cap.refp,
                bbox: cap.bbox,
                resources: form_resources,
                fill: cap.fill,
                stroke: cap.stroke,
                balanced,
                image: None,
            },
        );
        self.sync_e(fr);
    }

    /// `pdf:uxobj @name`: a form or a named image at A.
    fn use_form(&mut self, fr: &mut Frame, name: &str, ti: &Ti, a: (f64, f64)) {
        let Some(form) = self.forms.get(name) else {
            self.flag("unknown form XObject");
            return;
        };
        if let Some((img, img_ti)) = form.image.clone() {
            let mut t = if ti.width.is_some() || ti.has_height() || ti.bbox.is_some() { ti.clone() } else { img_ti.clone() };
            t.hide = false;
            if !ti.matrix.is_identity() {
                t.matrix = ti.matrix;
            }
            self.draw_image(fr, &img, &t, a);
            return;
        }
        let fit = fit_form(ti, form.bbox);
        // E group: K·Tr(A)·ti·fit·Tr(−ref)·K (the form's own SVG space is K of its device space)
        let t = K.mul(&Mat::translate(a.0, a.1)).mul(&ti.matrix).mul(&fit).mul(&Mat::translate(-form.refp.0, -form.refp.1)).mul(&K);
        self.sync_e(fr);
        self.replay(fr, name, &t, false);
    }

    /// `/Name Do` in PDF code: drawn at the CTM (no move to the current point).
    fn do_xobject(&mut self, fr: &mut Frame, name: &str) {
        let target = self.lookup_resource_ref(fr, "XObject", name);
        let Some(target) = target else {
            self.flag("unknown XObject");
            return;
        };
        let Some(form) = self.forms.get(&target) else {
            self.flag("unknown XObject");
            return;
        };
        if let Some((img, img_ti)) = form.image.clone() {
            // the unit square at the CTM
            let mut t = img_ti.clone();
            t.matrix = Mat::ID;
            let _ = t;
            let place = fr.user_to_e();
            let flip = Mat::new(1.0, 0.0, 0.0, -1.0, 0.0, 1.0);
            let content = match &img.source {
                ImageSource::Raster { .. } => Mat::ID,
                ImageSource::Box { llx, lly, urx, ury } => Mat::new(urx - llx, 0.0, 0.0, ury - lly, *llx, *lly),
            };
            self.sync_e(fr);
            self.raw(&format!(
                "<image x='0' y='0' width='1' height='1' preserveAspectRatio='none' transform='{}' xlink:href='{}'/>",
                place.mul(&content).mul(&flip).svg(),
                xml_escape(&img.href)
            ));
            return;
        }
        let t = fr.user_to_e().mul(&Mat::translate(-form.refp.0, -form.refp.1)).mul(&K);
        self.sync_e(fr);
        self.replay(fr, &target, &t, false);
    }

    fn lookup_resource_ref(&self, fr: &Frame, category: &str, name: &str) -> Option<String> {
        let find = |o: &Obj| -> Option<String> {
            match o.get(name) {
                Some(Obj::Ref(r)) => Some(r.clone()),
                _ => None,
            }
        };
        if let Some(cat) = fr.resources.get(category).and_then(|c| self.lookup(c, 0)) {
            if let Some(r) = find(&cat) {
                return Some(r);
            }
        }
        for v in self.objects.values() {
            if let Some(r) = find(v) {
                return Some(r);
            }
        }
        if self.forms.contains_key(name) {
            return Some(name.to_string());
        }
        None
    }

    /// The soft mask of `gs`: the form drawn into an SVG mask (luminosity), applied until the state is restored.
    fn soft_mask(&mut self, fr: &mut Frame, form_name: &str) {
        let Some(form) = self.forms.get(form_name) else {
            self.flag("unknown soft mask");
            return;
        };
        // the mask is used at the clip-group level: G · (E-group placement)
        let t = fr.g().mul(&fr.user_to_e()).mul(&Mat::translate(-form.refp.0, -form.refp.1)).mul(&K);
        let id = self.id();
        self.close_e(fr);
        self.raw(&format!("<mask id='{id}' maskUnits='userSpaceOnUse' x='-30000' y='-30000' width='60000' height='60000'>"));
        self.replay(fr, form_name, &t, true);
        self.raw("</mask>");
        self.push_group(fr, "", &format!("<g mask='url(#{id})'>"));
    }

    /// A captured form drawn again: its operations at its definition's DVI position, inside a group with
    /// transform `t` and clipped to its box.
    fn replay(&mut self, fr: &mut Frame, name: &str, t: &Mat, in_mask: bool) {
        if fr.depth >= MAX_DEPTH {
            self.flag("forms nested too deep");
            return;
        }
        let Some(form) = self.forms.get(name) else { return };
        let bytes = form.bytes.clone();
        let regs = form.regs;
        let coord = form.coord;
        let refp = form.refp;
        let bbox = form.bbox;
        let resources = form.resources.clone();
        let (fill, stroke) = (form.fill, form.stroke);
        let balanced = form.balanced;
        let fm_def = K;
        // the form's box in its own S space
        let corners = [(bbox[0], bbox[1]), (bbox[2], bbox[1]), (bbox[2], bbox[3]), (bbox[0], bbox[3])];
        let d: Vec<String> = corners
            .iter()
            .enumerate()
            .map(|(i, (x, y))| {
                let p = fm_def.apply(refp.0 + x, refp.1 + y);
                format!("{}{} {}", if i == 0 { 'M' } else { 'L' }, num(p.0), num(p.1))
            })
            .collect();
        let id = self.id();
        let _ = in_mask;
        self.raw(&format!("<g transform='{}'><clipPath id='{id}'><path d='{}Z'/></clipPath><g clip-path='url(#{id})'>", t.svg(), d.join("")));
        let mut sub = Frame::new(None, fill, stroke, regs, vec![coord], resources, fr.depth + 1);
        sub.hidden = fr.hidden;
        // DVI: to the definition's position and font, the operations, back
        let cur = fr.regs;
        let mut ops: Vec<u8> = vec![141];
        dvi::op_right(&mut ops, (regs.h - cur.h).round() as i32);
        dvi::op_down(&mut ops, (regs.v - cur.v).round() as i32);
        set_reg(&mut ops, 151, regs.w);
        set_reg(&mut ops, 156, regs.x);
        set_reg_v(&mut ops, 165, regs.y);
        set_reg_v(&mut ops, 170, regs.z);
        if self.dvi.fonts.contains_key(&regs.f) {
            dvi::op_fnt(&mut ops, regs.f);
        }
        self.out.extend_from_slice(&ops);
        self.sync_e(&mut sub);
        let body = if balanced { bytes } else { only_specials(&bytes) };
        if !balanced {
            self.flag("form XObject with unbalanced boxes");
        }
        self.run(&mut sub, &body);
        self.finish_frame(&mut sub);
        self.out.push(142);
        if self.dvi.fonts.contains_key(&cur.f) {
            dvi::op_fnt(&mut self.out, cur.f);
        }
        self.raw("</g></g>");
    }
}

/// dvipdfmx `scale_to_fit_F`: a form's box `b` (relative to its reference point) placed by the options.
fn fit_form(ti: &Ti, b: [f64; 4]) -> Mat {
    let (wd0, ht0) = ((b[2] - b[0]).max(1e-9), (b[3] - b[1]).max(1e-9));
    let (dx, dy) = (-b[0], -b[1]);
    let _ = (dx, dy);
    let (sx, sy, dp) = match (ti.width, ti.has_height()) {
        (Some(w), true) => (w / wd0, (ti.height.unwrap() + ti.depth) / ht0, ti.depth),
        (Some(w), false) => (w / wd0, w / wd0, 0.0),
        (None, true) => {
            let s = (ti.height.unwrap() + ti.depth) / ht0;
            (s, s, ti.depth)
        }
        (None, false) => (1.0, 1.0, 0.0),
    };
    if ti.bbox.is_some() {
        Mat::new(sx, 0.0, 0.0, sy, sx * -b[0], sy * -b[1] - dp)
    } else {
        Mat::new(sx, 0.0, 0.0, sy, 0.0, -dp)
    }
}

/// `w4`/`x4` set a spacing register (and move h by it): moved back at once.
fn set_reg(out: &mut Vec<u8>, op: u8, value: f64) {
    let v = value.round() as i32;
    out.push(op);
    dvi::put(out, v as u32, 4);
    dvi::op_right(out, -v);
}
fn set_reg_v(out: &mut Vec<u8>, op: u8, value: f64) {
    let v = value.round() as i32;
    out.push(op);
    dvi::put(out, v as u32, 4);
    dvi::op_down(out, -v);
}

fn balanced(bytes: &[u8]) -> bool {
    let mut r = Reader::new(bytes);
    let mut depth = 0i32;
    while r.pos < bytes.len() {
        match r.op() {
            Ok(Op::Push) => depth += 1,
            Ok(Op::Pop) => {
                depth -= 1;
                if depth < 0 {
                    return false;
                }
            }
            Ok(_) => {}
            Err(_) => return false,
        }
    }
    depth == 0
}

fn only_specials(bytes: &[u8]) -> Vec<u8> {
    let mut r = Reader::new(bytes);
    let mut out = vec![];
    while r.pos < bytes.len() {
        let s = r.pos;
        match r.op() {
            Ok(Op::Xxx(_)) => out.extend_from_slice(&bytes[s..r.pos]),
            Ok(_) => {}
            Err(_) => break,
        }
    }
    out
}

fn strip_pdf(s: &str) -> Option<&str> {
    s.strip_prefix("pdf:").map(|r| r.trim_start())
}

fn blend_css(name: &str) -> String {
    let mut out = String::new();
    for (i, c) in name.chars().enumerate() {
        if c.is_ascii_uppercase() && i > 0 {
            out.push('-');
        }
        out.push(c.to_ascii_lowercase());
    }
    out
}

/// `pdf:pagesize width W height H` / `papersize=W,H`, in bp.
pub fn page_size_special(s: &str) -> Option<(f64, f64)> {
    if let Some(rest) = s.strip_prefix("papersize=") {
        let mut it = rest.split(',');
        let mut w = Lexer::new(it.next()?.trim().as_bytes());
        let mut h = Lexer::new(it.next()?.trim().as_bytes());
        return Some((w.length()?, h.length()?));
    }
    let rest = strip_pdf(s)?.strip_prefix("pagesize")?;
    let mut lx = Lexer::new(rest.as_bytes());
    let (mut w, mut h) = (None, None);
    while let Some(key) = lx.ident() {
        match key.as_str() {
            "width" => w = lx.length(),
            "height" => h = lx.length(),
            "default" => {}
            _ => {
                lx.length();
            }
        }
    }
    Some((w?, h?))
}

/// A dvips colour (`rgb 1 0 0`, `cmyk …`, `gray 0.5`, `hsb …`, a name).
pub fn parse_color_spec(s: &str) -> Option<Rgb> {
    let mut parts = s.split_whitespace();
    let model = parts.next()?;
    let v: Vec<f64> = parts.filter_map(|p| p.parse().ok()).collect();
    Some(match model {
        "rgb" if v.len() >= 3 => Rgb(v[0], v[1], v[2]),
        "cmyk" if v.len() >= 4 => Rgb::cmyk(v[0], v[1], v[2], v[3]),
        "gray" | "grey" if !v.is_empty() => Rgb::gray(v[0]),
        "hsb" if v.len() >= 3 => {
            let (h, s, b) = (v[0], v[1], v[2]);
            let i = (h * 6.0).floor();
            let f = h * 6.0 - i;
            let (p, q, t) = (b * (1.0 - s), b * (1.0 - s * f), b * (1.0 - s * (1.0 - f)));
            match (i as i32).rem_euclid(6) {
                0 => Rgb(b, t, p),
                1 => Rgb(q, b, p),
                2 => Rgb(p, b, t),
                3 => Rgb(p, q, b),
                4 => Rgb(t, p, b),
                _ => Rgb(b, p, q),
            }
        }
        name => named_color(name)?,
    })
}

fn named_color(name: &str) -> Option<Rgb> {
    Some(match name {
        "Black" => Rgb::BLACK,
        "White" => Rgb(1.0, 1.0, 1.0),
        "Red" => Rgb::cmyk(0.0, 1.0, 1.0, 0.0),
        "Green" => Rgb::cmyk(1.0, 0.0, 1.0, 0.0),
        "Blue" => Rgb::cmyk(1.0, 1.0, 0.0, 0.0),
        "Cyan" => Rgb::cmyk(1.0, 0.0, 0.0, 0.0),
        "Magenta" => Rgb::cmyk(0.0, 1.0, 0.0, 0.0),
        "Yellow" => Rgb::cmyk(0.0, 0.0, 1.0, 0.0),
        "Gray" => Rgb::gray(0.5),
        _ => return None,
    })
}

/// `pdf:bcolor` arguments: fill (and stroke) colours as arrays or numbers.
fn parse_pdf_colors(s: &str) -> Option<(Rgb, Rgb)> {
    parse_pdf_colors_with(s, (Rgb::BLACK, Rgb::BLACK))
}

fn parse_pdf_colors_with(s: &str, current: (Rgb, Rgb)) -> Option<(Rgb, Rgb)> {
    let mut lx = Lexer::new(s.as_bytes());
    let read = |lx: &mut Lexer| -> Option<Rgb> {
        match lx.object()? {
            Obj::Array(a) => Some(Rgb::from_components(&a.iter().filter_map(|o| o.num()).collect::<Vec<_>>())),
            Obj::Num(v) => {
                let mut vals = vec![v];
                while let Some(n) = lx.number() {
                    vals.push(n);
                }
                Some(Rgb::from_components(&vals))
            }
            _ => None,
        }
    };
    lx.skip_white();
    if lx.rest().starts_with(b"fill") || lx.rest().starts_with(b"stroke") {
        let (mut f, mut st) = (current.1, current.0);
        while let Some(key) = lx.ident() {
            let c = read(&mut lx)?;
            if key == "fill" {
                f = c;
            } else {
                st = c;
            }
        }
        return Some((f, st));
    }
    let fill = read(&mut lx)?;
    let stroke = read(&mut lx).unwrap_or(fill);
    Some((fill, stroke))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn colors() {
        assert_eq!(parse_color_spec("rgb 1 0 0").unwrap().hex(), "#ff0000");
        assert_eq!(parse_color_spec("gray 0.5").unwrap().hex(), "#808080");
        assert_eq!(parse_color_spec("cmyk 0 1 1 0").unwrap().hex(), "#ff0000");
        let (f, s) = parse_pdf_colors("[1 0 0] [0 0 1]").unwrap();
        assert_eq!((f.hex().as_str(), s.hex().as_str()), ("#ff0000", "#0000ff"));
        let (f, _) = parse_pdf_colors("[0.5]").unwrap();
        assert_eq!(f.hex(), "#808080");
    }

    #[test]
    fn page_sizes() {
        let (w, h) = page_size_special("pdf:pagesize width 614.295pt height 794.96999pt").unwrap();
        assert!((w - 612.0).abs() < 0.01 && (h - 792.0).abs() < 0.01);
        let (w, _) = page_size_special("papersize=597.50787pt,845.04684pt").unwrap();
        assert!((w - 595.276).abs() < 0.01);
    }

    #[test]
    fn options() {
        let mut lx = Lexer::new(b"bbox 0 0 431.99895 287.99928 clip 0 width 85.35826pt (scatter.jpg)");
        let ti = read_ti(&mut lx);
        assert!(!ti.clip);
        assert!((ti.width.unwrap() - 85.35826 * 72.0 / 72.27).abs() < 1e-6);
        assert_eq!(lx.object(), Some(Obj::Str(b"scatter.jpg".to_vec())));
        let mut lx = Lexer::new(b"rotate 90");
        let ti = read_ti(&mut lx);
        assert!((ti.matrix.b - 1.0).abs() < 1e-9 && (ti.matrix.c + 1.0).abs() < 1e-9);
    }
}
