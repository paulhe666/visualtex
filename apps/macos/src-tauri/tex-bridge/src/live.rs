//! Live paragraphs: while the document's TeX run (0.5–1.5 s) makes the pages anew, the paragraph being edited
//! is typeset at once by a LuaLaTeX run that waits with the document's preamble read (lua/vtxlive.lua), and its
//! glyphs come back with their outlines (outlines.rs) for the page to draw over the paragraph's old lines.
//! LuaLaTeX for every engine: its paragraphs are TeX's own (the same line breaking, the same fonts read through
//! the map file or fontspec), so a paragraph comes out where the document's engine puts it — the page checks
//! that on the paragraph's text as last typeset before it shows any (the front end's verification).
//!
//! One server per document, in the folder TeX runs in, with its own job name (it never writes the document's
//! files): started after a compile when there is none for the current preamble, told to read the labels again
//! after each compile.

use std::collections::{HashMap, HashSet};
use std::io::{BufRead, BufReader, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, ChildStdin, Command, Stdio};
use std::sync::mpsc::{channel, Receiver, RecvTimeoutError};
use std::time::{Duration, Instant};

use serde::{Deserialize, Serialize};

use crate::engine::{self, Engine};
use crate::locate::TexDist;
use crate::outlines::{FontKey, Outlines};

pub const LIVE_JOB: &str = "_vtx_live";
const LUA: &str = include_str!("../lua/vtxlive.lua");
const LUA_FILE: &str = "_vtx_live.lua";
const MARK: &str = "@@VTXLIVE@@ ";
/// A paragraph that takes longer is not a paragraph TeX can finish alone (an unclosed group reading on).
const ANSWER_LIMIT: Duration = Duration::from_millis(2500);
/// The server is renewed after this many paragraphs (its log grows with each).
const RENEW_AFTER: u32 = 3000;
const SP_PER_BP: f64 = 65781.76;

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
    /// "unavailable" (no server: the preamble does not run on LuaLaTeX) | "failed"
    pub status: String,
    pub error: String,
    /// ms: TeX's, the whole call's
    pub tex_ms: f64,
    pub ms: f64,
    /// bp from the paragraph's top left (y down)
    pub height: f64,
    /// baseline, height, depth, width, shift (bp)
    pub lines: Vec<[f64; 5]>,
    /// font, x, y (baseline), code
    pub glyphs: Vec<(usize, f64, f64, u32)>,
    /// x, y (top), width, height
    pub rules: Vec<[f64; 4]>,
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

struct Server {
    child: Child,
    stdin: ChildStdin,
    answers: Receiver<String>,
    /// the preamble it read
    key: u64,
    ready: bool,
    asked: u32,
    /// TeX's fonts by its id, as the server described them
    fonts: HashMap<i64, LuaFont>,
}

impl Drop for Server {
    fn drop(&mut self) {
        let _ = self.stdin.write_all(b"Q\n");
        let _ = self.child.kill();
        let _ = self.child.wait();
    }
}

/// How the server is started (again, when a paragraph broke it).
struct Launch {
    dist: TexDist,
    dir: PathBuf,
    job: String,
    file: String,
    key: u64,
    shell: bool,
}

#[derive(Default)]
pub struct Live {
    server: Option<Server>,
    launch: Option<Launch>,
    /// a preamble whose server would not start (it does not run on LuaLaTeX)
    failed_key: Option<u64>,
    outlines: Outlines,
    sent: HashMap<String, HashSet<String>>,
    next: u64,
}

/// The live server's file: the main file's preamble as LuaLaTeX reads it (pdfLaTeX's default fonts, the
/// Computer Modern TFM fonts, when the document is a pdfLaTeX one), then the request loop.
fn server_file(engine: Engine, split: &engine::Split) -> String {
    let mut text = String::with_capacity(split.preamble.len() + 400);
    if engine == Engine::Pdflatex {
        text.push_str(concat!(
            "\\renewcommand\\encodingdefault{OT1}\\renewcommand\\rmdefault{cmr}\\renewcommand\\sfdefault{cmss}",
            "\\renewcommand\\ttdefault{cmtt}\\fontencoding{OT1}\\fontfamily{cmr}"
        ));
    }
    text.push_str(&split.preamble);
    text.push_str("\n\\begin{document}\n\\directlua{dofile(\"");
    text.push_str(LUA_FILE);
    text.push_str("\")}\n\\end{document}\n");
    text
}

impl Live {
    /// A server for this preamble, started now unless one runs or would not start. `job`: the document's job
    /// (its labels are read from a copy of its .aux).
    pub fn ensure(&mut self, dist: &TexDist, dir: &Path, job: &str, engine: Engine, split: Option<&engine::Split>, key: u64, shell: bool) {
        let Some(split) = split else {
            self.server = None;
            return;
        };
        let renew = self.server.as_ref().map(|s| s.asked >= RENEW_AFTER).unwrap_or(false);
        if let Some(server) = self.server.as_mut() {
            if server.key == key && !renew && matches!(server.child.try_wait(), Ok(None)) {
                // the labels of the run that just ended (a server still reading its preamble reads them once it is
                // waiting; the answer is skipped: paragraphs wait for their own)
                let _ = server.stdin.write_all(format!("R {job}.aux\n").as_bytes());
                let _ = server.stdin.flush();
                return;
            }
        }
        self.server = None;
        if self.failed_key == Some(key) || !dist.has("lualatex") {
            return;
        }
        self.launch = Some(Launch { dist: dist.clone(), dir: dir.to_path_buf(), job: job.to_string(), file: server_file(engine, split), key, shell });
        self.start();
    }

    fn start(&mut self) {
        self.server = None;
        let Some(l) = self.launch.as_ref() else { return };
        let dir = &l.dir;
        if std::fs::write(dir.join(LUA_FILE), LUA).is_err() || std::fs::write(dir.join(format!("{LIVE_JOB}.tex")), &l.file).is_err() {
            return;
        }
        let _ = std::fs::copy(dir.join(format!("{}.aux", l.job)), dir.join(format!("{LIVE_JOB}.aux")));
        let mut cmd = Command::new(l.dist.tool("lualatex"));
        cmd.arg("-interaction=nonstopmode")
            .arg(format!("-jobname={LIVE_JOB}"))
            .args(if l.shell { vec!["-shell-escape"] } else { vec![] })
            .arg(format!("{LIVE_JOB}.tex"))
            .current_dir(dir)
            .env("PATH", l.dist.path_env())
            .env("max_print_line", "10000")
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
                        let text = String::from_utf8_lossy(&line);
                        if let Some(json) = text.trim_end().strip_prefix(MARK) {
                            if tx.send(json.to_string()).is_err() {
                                break;
                            }
                        }
                    }
                }
            }
        });
        self.server = Some(Server { child, stdin, answers: rx, key: l.key, ready: false, asked: 0, fonts: HashMap::new() });
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
                Ok(json) => {
                    if serde_json::from_str::<LuaAnswer>(&json).map(|a| a.ready).unwrap_or(false) {
                        server.ready = true;
                        return true;
                    }
                }
                Err(RecvTimeoutError::Timeout) => return false,
                Err(RecvTimeoutError::Disconnected) => {
                    // it ended before its loop: the preamble does not run on LuaLaTeX
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
        let server = self.server.as_mut().unwrap();
        self.next += 1;
        let id = format!("q{}", self.next);
        let text = req.text.replace("\r\n", "\n");
        let head = format!("P {id} {} {} {} {} {}\n", req.hsize, req.indent, req.leftskip, req.rightskip, text.len());
        if server.stdin.write_all(head.as_bytes()).and_then(|_| server.stdin.write_all(text.as_bytes())).and_then(|_| server.stdin.flush()).is_err() {
            self.server = None;
            result.status = "failed".into();
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
            result.status = "failed".into();
            result.error = "the paragraph did not end".into();
            return result;
        };
        let a: LuaAnswer = match serde_json::from_str(&answer) {
            Ok(a) => a,
            Err(e) => {
                result.status = "failed".into();
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
            result.status = "failed".into();
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
        result.rules = a.rules.iter().map(|r| [bp(r[0]), bp(r[1]), bp(r[2]), bp(r[3])]).collect();
        result.glues = a.glues.iter().map(|g| [bp(g[0]), bp(g[1])]).collect();
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
            result.glyphs.push((index, bp(g[1]), bp(g[2]), code));
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
        result.ms = start.elapsed().as_secs_f64() * 1000.0;
        result
    }
}
