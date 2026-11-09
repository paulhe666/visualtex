//! Live paragraphs: while the document's TeX run (0.5–1.5 s) makes the pages anew, the paragraph being edited
//! is typeset at once by a run that waits with the document's preamble read, and drawn over the paragraph's old
//! lines (the front end checks first, on the paragraph's text as last typeset, that it comes out where the page
//! has it).
//!
//! Two kinds of server, one per document, in the folder TeX runs in, with their own job name (they never write
//! the document's files):
//! - the document's own engine (XeLaTeX, pdfLaTeX in DVI mode: lua/vtxlive-dvi.tex): the paragraph is shipped out
//!   as a page of the run's XDV/DVI, which becomes SVG as the pages do (screen copy, dvisvgm) — the same TeX, the
//!   same fonts and packages (xeCJK, microtype), so the same lines and glyphs;
//! - LuaLaTeX for LuaLaTeX documents (lua/vtxlive.lua): the paragraph's node list walked in Lua, its glyphs drawn
//!   from their outlines (outlines.rs).
//! Started after a compile when there is none for the current preamble; told to read the labels again after each.

use std::collections::{HashMap, HashSet};
use std::io::{BufRead, BufReader, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, ChildStdin, Command, Stdio};
use std::sync::mpsc::{channel, Receiver, RecvTimeoutError};
use std::time::{Duration, Instant};

use serde::{Deserialize, Serialize};

use crate::engine::{self, Engine};
use crate::images::{Drawers, ImageHost};
use crate::locate::TexDist;
use crate::outlines::{FontKey, Outlines};
use crate::svg::{self, SvgOptions};
use crate::tfm::Tfm;

pub const LIVE_JOB: &str = "_vtx_live";
const LUA: &str = include_str!("../lua/vtxlive.lua");
const LUA_FILE: &str = "_vtx_live.lua";
const DVI_LOOP: &str = include_str!("../lua/vtxlive-dvi.tex");
const DVI_LOOP_FILE: &str = "_vtx_live_loop.tex";
const PARAGRAPH_FILE: &str = "_vtx_live_par.tex";
const LABELS_FILE: &str = "_vtx_live_labels.tex";
const MARK: &str = "@@VTXLIVE@@ ";
/// A paragraph that takes longer is not a paragraph TeX can finish alone (an unclosed group reading on).
const ANSWER_LIMIT: Duration = Duration::from_millis(2500);
/// The server is renewed after this many paragraphs (its log grows with each; a DVI server's output too).
const RENEW_LUA: u32 = 3000;
const RENEW_DVI: u32 = 300;
const SP_PER_BP: f64 = 65781.76;
const PT_TO_BP: f64 = 72.0 / 72.27;

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LiveRequest {
    pub doc: String,
    /// the paragraph's LaTeX as the source has it
    pub text: String,
    /// sp: the line width, the indent (-1: the document's \parindent, 0: none), \leftskip, \rightskip
    pub hsize: i64,
    #[serde(default = "minus_one")]
    pub indent: i64,
    #[serde(default)]
    pub leftskip: i64,
    #[serde(default)]
    pub rightskip: i64,
    /// the reader: the paths it has been sent (a new one gets them all again)
    #[serde(default)]
    pub session: String,
}

fn minus_one() -> i64 {
    -1
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct LiveFont {
    /// index in this answer's glyphs
    pub index: usize,
    /// the outline set's name ("" when the font cannot be drawn)
    pub key: String,
    /// path units to bp: x, y (times the font's size)
    pub sx: f64,
    pub sy: f64,
    /// bp
    pub size: f64,
    /// x += slant × y; x ×= extend
    pub slant: f64,
    pub extend: f64,
}

#[derive(Clone, Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct LiveResult {
    /// "ok" | "error" (TeX's errors in the paragraph) | "migrated" (footnotes, floats: not alone) | "starting" |
    /// "unavailable" (no server: the preamble does not run) | "failed"
    pub status: String,
    pub error: String,
    /// ms: TeX's, the whole call's
    pub tex_ms: f64,
    pub ms: f64,
    /// bp from the paragraph's top left (y down)
    pub height: f64,
    /// baseline, height, depth, width, shift (bp)
    pub lines: Vec<[f64; 5]>,
    /// the paragraph as SVG (the document's own engine): its user coordinates are bp from the paragraph's top left
    pub svg: Option<String>,
    /// font, x, y (baseline), code, colour (index into `colors`; 0 black)
    pub glyphs: Vec<(usize, f64, f64, u32, usize)>,
    /// x, y (top), width, height, colour
    pub rules: Vec<[f64; 5]>,
    /// CSS colours of the glyphs and rules ("#000000" first)
    pub colors: Vec<String>,
    /// where each glue of a line ends: x, baseline y (as SyncTeX records glue: a page shown as a picture is
    /// checked against these)
    pub glues: Vec<[f64; 2]>,
    pub fonts: Vec<LiveFont>,
    /// "<key>:<code>" -> SVG path data (y up), the ones this reader has not been sent
    pub paths: HashMap<String, String>,
    /// a glyph has no outline here (a bitmap font): the paragraph cannot be drawn exactly
    pub undrawable: usize,
}

#[derive(Clone, Debug, Deserialize, Default)]
struct LuaFont {
    id: i64,
    kind: String,
    #[serde(default)]
    file: String,
    #[serde(default)]
    enc: String,
    #[serde(default)]
    sub: u32,
    #[serde(default)]
    size: i64,
    #[serde(default)]
    slant: i64,
    #[serde(default = "thousand")]
    extend: i64,
}

fn thousand() -> i64 {
    1000
}

#[derive(Debug, Deserialize, Default)]
struct LuaAnswer {
    #[serde(default)]
    ok: bool,
    #[serde(default)]
    ready: bool,
    #[serde(default)]
    ms: f64,
    #[serde(default)]
    height: i64,
    #[serde(default)]
    migrated: bool,
    #[serde(default)]
    errors: u32,
    #[serde(default)]
    error: String,
    #[serde(default)]
    lines: Vec<[i64; 5]>,
    #[serde(default)]
    glyphs: Vec<[i64; 4]>,
    #[serde(default)]
    rules: Vec<[i64; 4]>,
    #[serde(default)]
    glues: Vec<[i64; 2]>,
    #[serde(default)]
    fonts: Vec<LuaFont>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Kind {
    Lua,
    Dvi,
}

struct Server {
    child: Child,
    stdin: ChildStdin,
    /// Lua: the JSON answers; DVI: every line TeX writes to the terminal
    answers: Receiver<String>,
    kind: Kind,
    /// the preamble it read
    key: u64,
    ready: bool,
    asked: u32,
    /// TeX's fonts by its id, as the Lua server described them
    fonts: HashMap<i64, LuaFont>,
    /// DVI: TFM metrics and pictures for the screen copy of a paragraph's page, the fonts' files
    tfm: HashMap<String, Option<Tfm>>,
    images: Option<ImageHost>,
    files: crate::dvilive::FontFiles,
}

impl Drop for Server {
    fn drop(&mut self) {
        let _ = self.stdin.write_all(if self.kind == Kind::Lua { b"Q\n" } else { b"\\stop\n" });
        let _ = self.child.kill();
        let _ = self.child.wait();
    }
}

/// How the server is started (again, when a paragraph broke it).
struct Launch {
    dist: TexDist,
    dir: PathBuf,
    job: String,
    engine: Engine,
    file: String,
    key: u64,
    shell: bool,
    /// the folder for the paragraphs' screen copies and SVG
    work: PathBuf,
    svg_font_format: String,
    paper: (f64, f64),
}

#[derive(Default)]
pub struct Live {
    server: Option<Server>,
    launch: Option<Launch>,
    /// a preamble whose server would not start (it does not run on its engine here)
    failed_key: Option<u64>,
    outlines: Outlines,
    sent: HashMap<String, HashSet<String>>,
    next: u64,
}

fn kind_of(engine: Engine) -> Kind {
    if engine == Engine::Lualatex {
        Kind::Lua
    } else {
        Kind::Dvi
    }
}

/// The server's main file: the main file's preamble as the full run reads it, then the request loop.
fn server_file(engine: Engine, split: &engine::Split) -> String {
    let mut text = String::with_capacity(split.preamble.len() + 400);
    match kind_of(engine) {
        Kind::Lua => {
            text.push_str(&split.preamble);
            text.push_str("\n\\begin{document}\n\\directlua{dofile(\"");
            text.push_str(LUA_FILE);
            text.push_str("\")}\n\\end{document}\n");
        }
        Kind::Dvi => {
            text.push_str(engine.injection());
            text.push_str(&split.preamble);
            text.push_str("\n\\begin{document}\n\\input{");
            text.push_str(DVI_LOOP_FILE);
            text.push_str("}\n");
        }
    }
    text
}

/// The \newlabel and \bibcite lines of an .aux (what a paragraph's \ref and \cite need).
fn label_lines(aux: &Path) -> String {
    let text = std::fs::read_to_string(aux).unwrap_or_default();
    let mut out = String::new();
    for line in text.lines() {
        if line.starts_with("\\newlabel{") || line.starts_with("\\bibcite{") {
            out.push_str(line);
            out.push('\n');
        }
    }
    out
}

/// \showbox of the paragraph's vbox (depth 1): its lines (baseline, height, depth, width, shift; bp from the top),
/// its height and whether material left the lines (inserts: footnotes, floats).
fn box_lines(shown: &[String]) -> (Vec<[f64; 5]>, f64, bool) {
    let num = |s: &str| -> f64 {
        let t: String = s.trim_start().chars().take_while(|c| c.is_ascii_digit() || *c == '.' || *c == '-').collect();
        t.parse::<f64>().unwrap_or(0.0)
    };
    // "(H+D)xW"
    let dims = |s: &str| -> Option<(f64, f64, f64)> {
        let open = s.find('(')?;
        let plus = s[open..].find('+')? + open;
        let close = s[plus..].find(')')? + plus;
        let x = s[close..].find('x')? + close;
        Some((num(&s[open + 1..plus]), num(&s[plus + 1..close]), num(&s[x + 1..])))
    };
    let mut lines = vec![];
    let mut y = 0.0;
    let mut migrated = false;
    let mut started = false;
    for l in shown {
        if !started {
            started = l.starts_with("> \\box");
            continue;
        }
        let Some(item) = l.strip_prefix('.') else {
            if l.starts_with('\\') {
                continue; // the box itself
            }
            break;
        };
        if item.starts_with('.') {
            continue; // deeper
        }
        if let Some(rest) = item.strip_prefix("\\hbox") {
            if let Some((h, d, w)) = dims(rest) {
                y += h;
                let shift = rest.find("shifted ").map(|i| num(&rest[i + 8..])).unwrap_or(0.0);
                lines.push([y * PT_TO_BP, h * PT_TO_BP, d * PT_TO_BP, w * PT_TO_BP, shift * PT_TO_BP]);
                y += d;
            }
        } else if let Some(rest) = item.strip_prefix("\\vbox").or_else(|| item.strip_prefix("\\rule")) {
            if let Some((h, d, _)) = dims(rest) {
                y += h + d;
            }
        } else if let Some(rest) = item.strip_prefix("\\glue") {
            // "\glue(\baselineskip) 5.8", "\glue 3.0 plus 1.0"
            let rest = if rest.starts_with('(') { rest.find(')').map(|i| &rest[i + 1..]).unwrap_or(rest) } else { rest };
            y += num(rest);
        } else if let Some(rest) = item.strip_prefix("\\kern") {
            y += num(rest);
        } else if item.starts_with("\\insert") || item.starts_with("\\vadjust") {
            migrated = true;
        }
    }
    (lines, y * PT_TO_BP, migrated)
}

impl Live {
    /// A server for this preamble, started now unless one runs or would not start. `job`: the document's job
    /// (its labels are read from a copy of its .aux); `work`: where a paragraph's page is made into SVG.
    #[allow(clippy::too_many_arguments)]
    pub fn ensure(
        &mut self,
        dist: &TexDist,
        dir: &Path,
        job: &str,
        engine: Engine,
        split: Option<&engine::Split>,
        key: u64,
        shell: bool,
        work: &Path,
        svg_font_format: &str,
        paper: (f64, f64),
    ) {
        let Some(split) = split else {
            self.server = None;
            return;
        };
        let kind = kind_of(engine);
        let renew = self.server.as_ref().map(|s| s.asked >= if s.kind == Kind::Lua { RENEW_LUA } else { RENEW_DVI }).unwrap_or(false);
        if let Some(server) = self.server.as_mut() {
            if server.key == key && server.kind == kind && !renew && matches!(server.child.try_wait(), Ok(None)) {
                // the labels of the run that just ended (a server still reading its preamble reads them once it is
                // waiting; the answer is skipped: paragraphs wait for their own)
                match kind {
                    Kind::Lua => {
                        let _ = server.stdin.write_all(format!("R {job}.aux\n").as_bytes());
                    }
                    Kind::Dvi => {
                        if std::fs::write(dir.join(LABELS_FILE), label_lines(&dir.join(format!("{job}.aux")))).is_ok() {
                            let _ = server.stdin.write_all(format!("\\vtx@labels{{{LABELS_FILE}}}\n").as_bytes());
                        }
                    }
                }
                let _ = server.stdin.flush();
                return;
            }
        }
        self.server = None;
        let tool = match kind {
            Kind::Lua => "lualatex",
            Kind::Dvi => engine.name(),
        };
        if self.failed_key == Some(key) || !dist.has(tool) || (kind == Kind::Dvi && !dist.dvisvgm) {
            return;
        }
        self.launch = Some(Launch {
            dist: dist.clone(),
            dir: dir.to_path_buf(),
            job: job.to_string(),
            engine,
            file: server_file(engine, split),
            key,
            shell,
            work: work.to_path_buf(),
            svg_font_format: svg_font_format.to_string(),
            paper,
        });
        self.start();
    }

    fn start(&mut self) {
        self.server = None;
        let Some(l) = self.launch.as_ref() else { return };
        let kind = kind_of(l.engine);
        let dir = &l.dir;
        let extra = match kind {
            Kind::Lua => std::fs::write(dir.join(LUA_FILE), LUA),
            Kind::Dvi => std::fs::write(dir.join(DVI_LOOP_FILE), DVI_LOOP),
        };
        if extra.is_err() || std::fs::write(dir.join(format!("{LIVE_JOB}.tex")), &l.file).is_err() {
            return;
        }
        let _ = std::fs::copy(dir.join(format!("{}.aux", l.job)), dir.join(format!("{LIVE_JOB}.aux")));
        // a DVI server's output: each run starts its own
        for ext in ["dvi", "xdv"] {
            let _ = std::fs::remove_file(dir.join(format!("{LIVE_JOB}.{ext}")));
        }
        let mut cmd = match kind {
            Kind::Lua => {
                let mut c = Command::new(l.dist.tool("lualatex"));
                c.arg("-interaction=nonstopmode");
                c
            }
            Kind::Dvi => {
                let mut c = Command::new(l.dist.tool(l.engine.name()));
                c.arg(if l.engine == Engine::Xelatex { "-no-pdf" } else { "-output-format=dvi" }).arg("-interaction=scrollmode");
                c
            }
        };
        cmd.arg(format!("-jobname={LIVE_JOB}"))
            .args(if l.shell { vec!["-shell-escape"] } else { vec![] })
            .arg(format!("{LIVE_JOB}.tex"))
            .current_dir(dir)
            .env("PATH", l.dist.path_env())
            .env("max_print_line", "10000")
            .env("error_line", "254")
            .env("half_error_line", "238")
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::null());
        crate::process::hide_window(&mut cmd);
        let Ok(mut child) = cmd.spawn() else { return };
        let (Some(stdin), Some(stdout)) = (child.stdin.take(), child.stdout.take()) else { return };
        let (tx, rx) = channel();
        std::thread::spawn(move || {
            let mut reader = BufReader::new(stdout);
            let mut line = Vec::new();
            loop {
                line.clear();
                match reader.read_until(b'\n', &mut line) {
                    Ok(0) | Err(_) => break,
                    Ok(_) => {
                        let text = String::from_utf8_lossy(&line).trim_end().to_string();
                        let out = match kind {
                            Kind::Lua => match text.strip_prefix(MARK) {
                                Some(json) => json.to_string(),
                                None => continue,
                            },
                            Kind::Dvi => text,
                        };
                        if tx.send(out).is_err() {
                            break;
                        }
                    }
                }
            }
        });
        let images = (kind == Kind::Dvi).then(|| {
            let drawers = Drawers { mutool: l.dist.mutool.clone(), gs: l.dist.gs.clone(), extractbb: Some(l.dist.tool("extractbb")).filter(|p| p.is_file()) };
            ImageHost::new(dir, &l.work.join("images"), drawers)
        });
        let files = crate::dvilive::FontFiles::new(Some(l.dist.tool("kpsewhich")).filter(|p| p.is_file()));
        self.server = Some(Server { child, stdin, answers: rx, kind, key: l.key, ready: false, asked: 0, fonts: HashMap::new(), tfm: HashMap::new(), images, files });
    }

    pub fn stop(&mut self) {
        self.server = None;
    }

    /// Waits (a little) for a server that is starting: true once it is ready.
    fn wait_ready(&mut self, limit: Duration) -> bool {
        let Some(server) = self.server.as_mut() else { return false };
        if server.ready {
            return true;
        }
        let start = Instant::now();
        loop {
            match server.answers.recv_timeout(limit.saturating_sub(start.elapsed())) {
                Ok(text) => {
                    let ready = match server.kind {
                        Kind::Lua => serde_json::from_str::<LuaAnswer>(&text).map(|a| a.ready).unwrap_or(false),
                        Kind::Dvi => text.starts_with("VTXREADY"),
                    };
                    if ready {
                        server.ready = true;
                        return true;
                    }
                }
                Err(RecvTimeoutError::Timeout) => return false,
                Err(RecvTimeoutError::Disconnected) => {
                    // it ended before its loop: the preamble does not run here
                    self.failed_key = Some(server.key);
                    self.server = None;
                    return false;
                }
            }
        }
    }

    /// Typesets one paragraph.
    pub fn paragraph(&mut self, req: &LiveRequest) -> LiveResult {
        let start = Instant::now();
        let mut result = LiveResult { status: "unavailable".into(), ..Default::default() };
        if self.server.is_none() {
            return result;
        }
        // a server that started a moment ago may need its first second
        if !self.wait_ready(Duration::from_millis(if self.server.as_ref().map(|s| s.asked == 0).unwrap_or(false) { 1500 } else { 30 })) {
            result.status = if self.server.is_some() { "starting".into() } else { "unavailable".into() };
            return result;
        }
        self.next += 1;
        let id = format!("q{}", self.next);
        let mut result = match self.server.as_ref().map(|s| s.kind) {
            Some(Kind::Dvi) => self.dvi_paragraph(req, &id),
            _ => self.lua_paragraph(req, &id),
        };
        result.ms = start.elapsed().as_secs_f64() * 1000.0;
        result
    }

    fn lua_paragraph(&mut self, req: &LiveRequest, id: &str) -> LiveResult {
        let start = Instant::now();
        let mut result = LiveResult { status: "failed".into(), ..Default::default() };
        let server = self.server.as_mut().unwrap();
        let text = req.text.replace("\r\n", "\n");
        let head = format!("P {id} {} {} {} {} {}\n", req.hsize, req.indent, req.leftskip, req.rightskip, text.len());
        if server.stdin.write_all(head.as_bytes()).and_then(|_| server.stdin.write_all(text.as_bytes())).and_then(|_| server.stdin.flush()).is_err() {
            self.server = None;
            return result;
        }
        server.asked += 1;
        let answer = loop {
            match server.answers.recv_timeout(ANSWER_LIMIT.saturating_sub(start.elapsed())) {
                Ok(json) => {
                    if json.contains(&format!("\"id\":\"{id}\"")) {
                        break Some(json);
                    }
                }
                Err(_) => break None,
            }
        };
        let Some(answer) = answer else {
            // TeX reads on past the paragraph (an unclosed group, \end{document}): this server is done, a new one
            // starts
            self.start();
            result.error = "the paragraph did not end".into();
            return result;
        };
        let a: LuaAnswer = match serde_json::from_str(&answer) {
            Ok(a) => a,
            Err(e) => {
                result.error = e.to_string();
                return result;
            }
        };
        let server = self.server.as_mut().unwrap();
        for f in a.fonts {
            server.fonts.insert(f.id, f);
        }
        result.tex_ms = a.ms;
        if !a.ok {
            // a paragraph that broke the server's state (groups left open): a new one starts
            if a.error.contains("unbalanced") {
                self.start();
            }
            result.error = a.error;
            return result;
        }
        result.status = if a.errors > 0 {
            "error".into()
        } else if a.migrated {
            "migrated".into()
        } else {
            "ok".into()
        };
        result.error = a.error;
        let bp = |v: i64| v as f64 / SP_PER_BP;
        result.height = bp(a.height);
        result.lines = a.lines.iter().map(|l| [bp(l[0]), bp(l[1]), bp(l[2]), bp(l[3]), bp(l[4])]).collect();
        result.rules = a.rules.iter().map(|r| [bp(r[0]), bp(r[1]), bp(r[2]), bp(r[3]), 0.0]).collect();
        result.glues = a.glues.iter().map(|g| [bp(g[0]), bp(g[1])]).collect();
        result.colors = vec!["#000000".into()];
        let sent = self.sent.entry(req.session.clone()).or_default();
        let mut index_of: HashMap<i64, usize> = HashMap::new();
        for g in &a.glyphs {
            let (fid, code) = (g[0], g[3].max(0) as u32);
            let index = match index_of.get(&fid) {
                Some(&i) => i,
                None => {
                    let f = server.fonts.get(&fid).cloned().unwrap_or_default();
                    let key = (f.kind == "otf" || f.kind == "type1").then(|| FontKey { kind: f.kind.clone(), file: f.file.clone(), enc: f.enc.clone(), sub: f.sub });
                    let scale = key.as_ref().and_then(|k| self.outlines.scale(k));
                    let size = bp(f.size);
                    let (key_id, sx, sy) = match (&key, scale) {
                        (Some(k), Some((sx, sy))) => (k.id(), sx * size, sy * size),
                        _ => (String::new(), 0.0, 0.0),
                    };
                    let i = result.fonts.len();
                    result.fonts.push(LiveFont { index: i, key: key_id, sx, sy, size, slant: f.slant as f64 / 1000.0, extend: f.extend as f64 / 1000.0 });
                    index_of.insert(fid, i);
                    i
                }
            };
            result.glyphs.push((index, bp(g[1]), bp(g[2]), code, 0));
            let font = &result.fonts[index];
            if font.key.is_empty() {
                result.undrawable += 1;
                continue;
            }
            let name = format!("{}:{code}", font.key);
            if sent.contains(&name) {
                continue;
            }
            let f = server.fonts.get(&fid).cloned().unwrap_or_default();
            let key = FontKey { kind: f.kind, file: f.file, enc: f.enc, sub: f.sub };
            match self.outlines.path(&key, code) {
                Some(path) => {
                    result.paths.insert(name.clone(), path);
                    sent.insert(name);
                }
                None => result.undrawable += 1,
            }
        }
        result
    }

    fn dvi_paragraph(&mut self, req: &LiveRequest, id: &str) -> LiveResult {
        let start = Instant::now();
        let mut result = LiveResult { status: "failed".into(), ..Default::default() };
        let Some(l) = self.launch.as_ref() else { return result };
        let (dir, work, engine) = (l.dir.clone(), l.work.clone(), l.engine);
        let opts = SvgOptions { dvisvgm: l.dist.tool("dvisvgm"), font_format: l.svg_font_format.clone(), processes: 1 };
        let kpse = l.dist.tool("kpsewhich");
        let paper = l.paper;
        let job = l.job.clone();
        let server = self.server.as_mut().unwrap();
        if std::fs::write(dir.join(PARAGRAPH_FILE), req.text.replace("\r\n", "\n")).is_err() {
            return result;
        }
        let line = format!("\\vtx@para{{{id}}}{{{}}}{{{}}}{{{}}}{{{}}}{{{PARAGRAPH_FILE}}}\n", req.hsize, req.indent, req.leftskip, req.rightskip);
        if server.stdin.write_all(line.as_bytes()).and_then(|_| server.stdin.flush()).is_err() {
            self.server = None;
            return result;
        }
        server.asked += 1;
        // TeX's terminal lines up to the paragraph's end
        let mut shown = vec![];
        let done = format!("VTXDONE {id}");
        let broken = format!("VTXBROKEN {id}");
        let finished = loop {
            match server.answers.recv_timeout(ANSWER_LIMIT.saturating_sub(start.elapsed())) {
                Ok(text) => {
                    if text.ends_with(&done) {
                        break Some(true);
                    }
                    if text.ends_with(&broken) {
                        break Some(false);
                    }
                    shown.push(text);
                }
                Err(_) => break None,
            }
        };
        if finished != Some(true) {
            // the paragraph left a group open, or TeX reads on: this server is done, a new one starts
            self.start();
            result.error = "the paragraph did not end".into();
            return result;
        }
        result.tex_ms = start.elapsed().as_secs_f64() * 1000.0;
        let errors: Vec<&String> = shown.iter().filter(|l| l.starts_with("! ") && !l.starts_with("! OK")).collect();
        let (lines, height, migrated) = box_lines(&shown);
        result.lines = lines;
        result.height = height;
        result.status = if !errors.is_empty() {
            "error".into()
        } else if migrated {
            "migrated".into()
        } else {
            "ok".into()
        };
        result.error = errors.first().map(|e| e.trim_start_matches("! ").to_string()).unwrap_or_default();
        if result.status != "ok" || result.lines.is_empty() {
            return result;
        }
        // the paragraph's page in the run's output: a file of its own, then as the pages are shown
        let output = dir.join(format!("{LIVE_JOB}.{}", engine.output()));
        let Ok(bytes) = std::fs::read(&output) else {
            result.status = "failed".into();
            result.error = "no output".into();
            return result;
        };
        let Ok(mut dvi) = crate::dvi::Dvi::parse_partial(&bytes) else {
            result.status = "failed".into();
            result.error = "output not readable".into();
            return result;
        };
        let marker = format!("vtx:{id}");
        let Some(page) = dvi.pages.iter().rev().find(|p| p.body.windows(marker.len()).any(|w| w == marker.as_bytes())).cloned() else {
            result.status = "failed".into();
            result.error = "the paragraph's page is not out yet".into();
            return result;
        };
        // drawn from the fonts' outlines when everything on it can be (glyphs, rules, colours)
        let drawn = crate::dvilive::draw_page(&dvi, &page.body, &mut server.files);
        if drawn.undrawable.is_empty() {
            let sent = self.sent.entry(req.session.clone()).or_default();
            let mut undrawable = 0;
            for (key, size, slant, extend) in &drawn.fonts {
                let (sx, sy) = self.outlines.scale(key).unwrap_or((0.0, 0.0));
                let index = result.fonts.len();
                let ok = sx != 0.0;
                result.fonts.push(LiveFont { index, key: if ok { key.id() } else { String::new() }, sx: sx * size, sy: sy * size, size: *size, slant: *slant, extend: *extend });
            }
            for &(fi, x, y, code, color) in &drawn.glyphs {
                result.glyphs.push((fi, x, y, code, color));
                let font = &result.fonts[fi];
                if font.key.is_empty() {
                    undrawable += 1;
                    continue;
                }
                let name = format!("{}:{code}", font.key);
                if sent.contains(&name) {
                    continue;
                }
                match self.outlines.path(&drawn.fonts[fi].0, code) {
                    Some(path) => {
                        result.paths.insert(name.clone(), path);
                        sent.insert(name);
                    }
                    None => undrawable += 1,
                }
            }
            result.rules = drawn.rules.iter().map(|(r, c)| [r[0], r[1], r[2], r[3], *c as f64]).collect();
            result.colors = drawn.colors.clone();
            result.undrawable = undrawable;
            if undrawable == 0 {
                return result;
            }
            // (a glyph without an outline here: the page goes through dvisvgm)
            result.glyphs.clear();
            result.fonts.clear();
            result.rules.clear();
            result.undrawable = 0;
        }
        // every font defined so far, before the page (it may use fonts an earlier page defined)
        let server = self.server.as_mut().unwrap();
        let defs: Vec<u8> = dvi.fonts.values().flat_map(|d| d.raw.clone()).collect();
        let body = page.body.clone();
        dvi.pages = vec![crate::dvi::Page { before: defs.clone(), ..page }];
        dvi.post_fonts = defs;
        let one = dvi.write_with(&[body]);
        let _ = std::fs::create_dir_all(&work);
        let screen = work.join("live.dvi");
        let fls = std::fs::read_to_string(dir.join(format!("{job}.fls"))).unwrap_or_default();
        let inputs = crate::fonts::recorded_inputs(&fls, &dir);
        let images = server.images.get_or_insert_with(|| ImageHost::new(&dir, &work.join("images"), Drawers { mutool: None, gs: None, extractbb: None }));
        let pages = match crate::screen::make_screen(&one, &screen, &dir, &inputs, Some(&kpse), &mut server.tfm, images, paper) {
            Ok(p) => p,
            Err(e) => {
                result.status = "failed".into();
                result.error = e;
                return result;
            }
        };
        let Some(meta) = pages.first() else {
            result.status = "failed".into();
            return result;
        };
        if !meta.unsupported.is_empty() {
            // drawn from the PDF only (shadings, some pictures): not here
            result.status = "failed".into();
            result.error = meta.unsupported.join(", ");
            return result;
        }
        let made = svg::convert(&opts, &dir, &screen, &[(1, meta.hash_value)], &work.join("live-svg"));
        match made.get(&1) {
            Some(svg) => result.svg = Some(svg.clone()),
            None => {
                result.status = "failed".into();
                result.error = "dvisvgm".into();
            }
        }
        result
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn box_lines_read_showbox() {
        let shown: Vec<String> = [
            "(./_vtx_live_par.tex) [1] [1]",
            "> \\box55=",
            "\\vbox(24.87294+2.10786)x345.0",
            ".\\hbox(8.43149+2.1711)x345.0, glue set 0.23718 []",
            ".\\penalty 300",
            ".\\glue(\\baselineskip) 5.83887",
            ".\\hbox(8.43149+2.10786)x345.0, glue set 309.62988fil []",
            "",
            "! OK.",
        ]
        .iter()
        .map(|s| s.to_string())
        .collect();
        let (lines, height, migrated) = box_lines(&shown);
        assert_eq!(lines.len(), 2);
        assert!((lines[0][0] - 8.43149 * PT_TO_BP).abs() < 1e-9);
        assert!((lines[1][0] - (8.43149 + 2.1711 + 5.83887 + 8.43149) * PT_TO_BP).abs() < 1e-9);
        assert!((height - (8.43149 + 2.1711 + 5.83887 + 8.43149 + 2.10786) * PT_TO_BP).abs() < 1e-9);
        assert!(!migrated);
    }
}
