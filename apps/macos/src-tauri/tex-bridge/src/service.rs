//! The compile service behind the `tex_*` commands: one per app. Each LaTeX document has a build folder in the
//! app's cache (`work/`: the project's copy where TeX runs; `builds/<n>/`: what one run produced), a warm engine
//! waiting with the preamble read, and the SVG pages already made (by page hash).
//!
//! A compile: copy what changed into `work/`, run the engine (the warm one when the preamble is unchanged),
//! keep its output as build n, start the next warm engine, write the screen copy of the DVI/XDV and convert the
//! pages asked for. A newer compile of the same document kills a running one (the older call answers
//! "superseded"). Reruns (references, contents, BibTeX/Biber) are reported (`needsRerun`) for the caller to
//! ask for, so the first pass shows at once; `full` reruns until the document is settled (for export).

use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, ChildStdin, Command, Stdio};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Condvar, Mutex};
use std::time::{Duration, Instant};

use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine as _;
use serde::{Deserialize, Serialize};

use crate::engine::{self, Engine};
use crate::images::{Drawers, ImageHost};
use crate::locate::{self, TexDist};
use crate::process::{self, Killer};
use crate::project::{self, Manifest, ProjectFile};
use crate::screen::{self, PageMeta};
use crate::svg::{self, SvgOptions};
use crate::tfm::Tfm;
use crate::util::{hash_bytes, Fnv};

// ------------------------------------------------------------------------------------------- requests --

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CompileRequest {
    /// the document's id (its library id)
    pub doc: String,
    /// the project folder
    pub root: String,
    /// the main file, relative to the folder
    pub main: String,
    /// "xelatex" | "pdflatex" | "lualatex"; empty or "auto": chosen from the document
    #[serde(default)]
    pub engine: Option<String>,
    #[serde(default)]
    pub shell_escape: bool,
    /// editor text not saved yet, by relative path
    #[serde(default)]
    pub overlays: HashMap<String, String>,
    /// pages (1-based) whose SVG the answer should carry
    #[serde(default)]
    pub visible: Vec<usize>,
    /// rerun (and run BibTeX/Biber) until the document is settled
    #[serde(default)]
    pub full: bool,
    /// names this compile for `tex_preview` (its visible pages as soon as TeX has written them)
    #[serde(default)]
    pub ticket: Option<String>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PreviewRequest {
    pub doc: String,
    pub ticket: String,
}

/// The visible pages of a compile still running ("early"), or nothing more to come ("done").
#[derive(Clone, Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct PreviewResult {
    pub status: String,
    pub pages: Vec<PageMeta>,
    pub svgs: Vec<SvgPage>,
    /// ms from the compile's start
    pub at: u64,
}

#[derive(Clone, Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Timings {
    pub sync: u64,
    pub tex: u64,
    pub bib: u64,
    pub screen: u64,
    pub svg: u64,
    pub total: u64,
    pub passes: u32,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SvgPage {
    pub page: usize,
    pub hash: String,
    pub svg: String,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CompileResult {
    /// "ok" | "failed" (no pages) | "superseded" | "notex"
    pub status: String,
    pub build: u64,
    pub engine: String,
    /// why this engine ("% !TEX program", "fontspec", …)
    pub engine_reason: String,
    /// "svg": pages from the DVI/XDV; "raster": pictures of the PDF (lualatex)
    pub page_mode: String,
    pub pages: Vec<PageMeta>,
    pub svgs: Vec<SvgPage>,
    pub log: String,
    pub needs_rerun: bool,
    pub rerun_reason: String,
    /// the engine had the preamble read already
    pub warm: bool,
    pub timings: Timings,
    pub error: String,
}

impl CompileResult {
    fn empty(status: &str) -> CompileResult {
        CompileResult {
            status: status.into(),
            build: 0,
            engine: String::new(),
            engine_reason: String::new(),
            page_mode: "svg".into(),
            pages: vec![],
            svgs: vec![],
            log: String::new(),
            needs_rerun: false,
            rerun_reason: String::new(),
            warm: false,
            timings: Timings::default(),
            error: String::new(),
        }
    }
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PagesRequest {
    pub doc: String,
    pub build: u64,
    pub pages: Vec<usize>,
    /// pixels per bp for pages shown as pictures
    #[serde(default)]
    pub scale: Option<f64>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PageOut {
    pub page: usize,
    pub hash: String,
    /// inline SVG
    pub svg: Option<String>,
    /// a PNG of the PDF page (base64): pages the SVG cannot show, PDF-only engines
    pub png: Option<String>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BuildRequest {
    pub doc: String,
    pub build: u64,
    /// a path to copy the PDF to (export)
    #[serde(default)]
    pub dest: Option<String>,
    /// answer the PDF's bytes (base64)
    #[serde(default)]
    pub bytes: bool,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfOut {
    pub path: String,
    pub size: u64,
    pub base64: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SynctexOut {
    /// the .synctex.gz, base64
    pub base64: String,
    /// the folder TeX ran in (input paths in the SyncTeX file are under it)
    pub work: String,
    /// files of the run standing for the main file (the warm wrapper and body)
    pub aliases: Vec<String>,
    pub main: String,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FileRequest {
    pub root: String,
    #[serde(default)]
    pub path: String,
    /// text to write (UTF-8)
    #[serde(default)]
    pub text: Option<String>,
    /// bytes to write
    #[serde(default)]
    pub base64: Option<String>,
    /// for open: a preferred main file
    #[serde(default)]
    pub main: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectInfo {
    pub root: String,
    pub main: Option<String>,
    pub files: Vec<ProjectFile>,
    pub engine: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TexStatus {
    pub found: bool,
    pub dist: Option<TexDist>,
}

// ------------------------------------------------------------------------------------------- state --

struct Running {
    child: Child,
    stdin: Option<ChildStdin>,
    output: Arc<Mutex<Vec<u8>>>,
}

struct Warm {
    run: Running,
    key: u64,
}

struct Build {
    dir: PathBuf,
    /// the folder TeX ran in
    run_dir: PathBuf,
    job: String,
    engine: Engine,
    pages: Vec<PageMeta>,
    pdf: Option<PathBuf>,
}

#[derive(Default)]
struct RunState {
    manifest: Manifest,
    warm: Option<Warm>,
    tfm: HashMap<String, Option<Tfm>>,
    images: Option<ImageHost>,
    /// aux-like files after the last run (rerun detection)
    aux: HashMap<String, u64>,
    /// what BibTeX/Biber last ran on
    bib_key: u64,
    next_build: u64,
}

#[derive(Default)]
struct ViewState {
    builds: BTreeMap<u64, Arc<Build>>,
    svg: HashMap<u64, Arc<String>>,
    svg_order: Vec<u64>,
}

#[derive(Default)]
struct Progress {
    ticket: String,
    result: Option<PreviewResult>,
    done: bool,
}

/// A document nobody asked anything of for this long: its engines (the warm one, the live server) stop until it is
/// asked again (a document left open in a window that is gone keeps hundreds of MB in each).
const IDLE_ENGINES: Duration = Duration::from_secs(15 * 60);

struct Doc {
    dir: PathBuf,
    /// when it was last asked something (ms since the service started)
    touched: AtomicU64,
    live: Mutex<crate::live::Live>,
    generation: AtomicU64,
    killer: Killer,
    run: Mutex<RunState>,
    view: Mutex<ViewState>,
    progress: Mutex<Progress>,
    progress_cv: Condvar,
}

pub struct Service {
    started: Instant,
    cache_root: PathBuf,
    projects_root: PathBuf,
    docs: Mutex<HashMap<String, Arc<Doc>>>,
    roots: Mutex<HashSet<PathBuf>>,
    dist: Mutex<Option<Option<TexDist>>>,
    pub svg_font_format: String,
}

const KEEP_BUILDS: usize = 4;
const SVG_CACHE: usize = 400;
const TEX_LIMIT: Duration = Duration::from_secs(240);

fn ms(t: Instant) -> u64 {
    t.elapsed().as_millis() as u64
}

fn safe_id(id: &str) -> String {
    let s: String = id.chars().map(|c| if c.is_ascii_alphanumeric() || c == '-' || c == '_' { c } else { '_' }).take(64).collect();
    if s.is_empty() {
        "doc".into()
    } else {
        s
    }
}

impl Service {
    /// `cache_root`: build folders (the app's cache); `projects_root`: projects made in the app (new documents,
    /// opened archives).
    pub fn new(cache_root: PathBuf, projects_root: PathBuf) -> Service {
        Service {
            started: Instant::now(),
            cache_root,
            projects_root,
            docs: Mutex::new(HashMap::new()),
            roots: Mutex::new(HashSet::new()),
            dist: Mutex::new(None),
            svg_font_format: "woff".into(),
        }
    }

    pub fn dist(&self) -> Option<TexDist> {
        let mut d = self.dist.lock().unwrap();
        if d.is_none() {
            *d = Some(locate::locate());
        }
        d.clone().unwrap()
    }

    pub fn status(&self) -> TexStatus {
        // look again (TeX may have been installed since)
        *self.dist.lock().unwrap() = None;
        let dist = self.dist();
        TexStatus { found: dist.is_some(), dist }
    }

    fn doc(&self, id: &str) -> Arc<Doc> {
        let now = self.started.elapsed().as_millis() as u64;
        let mut docs = self.docs.lock().unwrap();
        // the engines of documents idle for long stop (not while a compile or a paragraph holds them)
        for (other, d) in docs.iter() {
            if other != id && now.saturating_sub(d.touched.load(Ordering::Relaxed)) > IDLE_ENGINES.as_millis() as u64 {
                if let Ok(mut run) = d.run.try_lock() {
                    if let Some(mut w) = run.warm.take() {
                        let _ = w.run.child.kill();
                        let _ = w.run.child.wait();
                    }
                }
                if let Ok(mut live) = d.live.try_lock() {
                    live.stop();
                }
            }
        }
        let doc = docs
            .entry(id.to_string())
            .or_insert_with(|| {
                Arc::new(Doc {
                    dir: self.cache_root.join(safe_id(id)),
                    touched: AtomicU64::new(now),
                    live: Mutex::new(crate::live::Live::default()),
                    generation: AtomicU64::new(0),
                    killer: Killer::default(),
                    run: Mutex::new(RunState::default()),
                    view: Mutex::new(ViewState::default()),
                    progress: Mutex::new(Progress::default()),
                    progress_cv: Condvar::new(),
                })
            })
            .clone();
        doc.touched.store(now, Ordering::Relaxed);
        doc
    }

    // ----------------------------------------------------------------------------------- project files --

    fn allowed_root(&self, root: &str) -> Result<PathBuf, String> {
        let p = PathBuf::from(root);
        if !p.is_absolute() {
            return Err("The project folder must be an absolute path".into());
        }
        let canon = p.canonicalize().map_err(|e| format!("{root}: {e}"))?;
        let projects = self.projects_root.canonicalize().unwrap_or(self.projects_root.clone());
        if canon.starts_with(&projects) || self.roots.lock().unwrap().contains(&canon) {
            Ok(canon)
        } else {
            Err("The folder was not opened as a project".into())
        }
    }

    /// Open a project: a folder, or a .tex file (its folder; the file the main one). The folder may then be read
    /// and written through the other calls.
    pub fn project_open(&self, req: &FileRequest) -> Result<ProjectInfo, String> {
        let given = PathBuf::from(&req.root);
        if !given.is_absolute() {
            return Err("An absolute path is needed".into());
        }
        let (root, file_main) = if given.is_file() {
            let parent = given.parent().ok_or("The file has no folder")?.to_path_buf();
            let name = given.file_name().map(|n| n.to_string_lossy().to_string());
            (parent, name)
        } else {
            (given, None)
        };
        let root = root.canonicalize().map_err(|e| format!("{}: {e}", root.display()))?;
        self.roots.lock().unwrap().insert(root.clone());
        let files = project::list_files(&root);
        let main = req.main.clone().filter(|m| files.iter().any(|f| &f.path == m)).or(file_main).or_else(|| project::find_main(&root, &files));
        let engine = main.as_ref().and_then(|m| std::fs::read(root.join(m)).ok()).map(|b| {
            let text = String::from_utf8_lossy(&b).to_string();
            engine::detect(&text, &self.class_texts(&root, &files)).0.name().to_string()
        });
        Ok(ProjectInfo { root: root.to_string_lossy().to_string(), main, files, engine })
    }

    fn class_texts(&self, root: &Path, files: &[ProjectFile]) -> Vec<String> {
        files
            .iter()
            .filter(|f| {
                let l = f.path.to_ascii_lowercase();
                l.ends_with(".cls") || l.ends_with(".sty")
            })
            .take(40)
            .filter_map(|f| std::fs::read(root.join(&f.path)).ok().map(|b| String::from_utf8_lossy(&b).to_string()))
            .collect()
    }

    /// A new, empty project folder in the app's projects folder.
    pub fn project_new(&self, name_hint: &str) -> Result<ProjectInfo, String> {
        std::fs::create_dir_all(&self.projects_root).map_err(|e| e.to_string())?;
        let stem: String = name_hint.chars().filter(|c| c.is_alphanumeric() || *c == '-' || *c == '_').take(40).collect();
        let stem = if stem.is_empty() { "latex".to_string() } else { stem };
        let mut n = 0;
        let root = loop {
            let candidate = self.projects_root.join(if n == 0 { stem.clone() } else { format!("{stem}-{n}") });
            if !candidate.exists() {
                break candidate;
            }
            n += 1;
        };
        std::fs::create_dir_all(&root).map_err(|e| e.to_string())?;
        let root = root.canonicalize().map_err(|e| e.to_string())?;
        Ok(ProjectInfo { root: root.to_string_lossy().to_string(), main: None, files: vec![], engine: None })
    }

    pub fn project_files(&self, req: &FileRequest) -> Result<ProjectInfo, String> {
        let root = self.allowed_root(&req.root)?;
        let files = project::list_files(&root);
        Ok(ProjectInfo { root: root.to_string_lossy().to_string(), main: None, files, engine: None })
    }

    pub fn project_read(&self, req: &FileRequest) -> Result<String, String> {
        let root = self.allowed_root(&req.root)?;
        let rel = project::safe_rel(&req.path).ok_or("Not a path inside the project")?;
        let bytes = std::fs::read(root.join(rel)).map_err(|e| format!("{}: {e}", req.path))?;
        Ok(BASE64.encode(bytes))
    }

    pub fn project_write(&self, req: &FileRequest) -> Result<(), String> {
        let root = self.allowed_root(&req.root)?;
        let rel = project::safe_rel(&req.path).ok_or("Not a path inside the project")?;
        let bytes = match (&req.text, &req.base64) {
            (Some(t), _) => t.as_bytes().to_vec(),
            (None, Some(b)) => BASE64.decode(b.trim()).map_err(|e| e.to_string())?,
            _ => return Err("Nothing to write".into()),
        };
        let dest = root.join(rel);
        if let Some(parent) = dest.parent() {
            std::fs::create_dir_all(parent).map_err(|e| e.to_string())?;
        }
        // write next to it, then replace: a crash never leaves half a file
        let tmp = dest.with_extension("vtx-saving");
        std::fs::write(&tmp, &bytes).map_err(|e| e.to_string())?;
        std::fs::rename(&tmp, &dest).map_err(|e| {
            let _ = std::fs::remove_file(&tmp);
            e.to_string()
        })
    }

    // ----------------------------------------------------------------------------------- compile --

    /// Stop the document's engines and forget its builds (the document was closed or deleted).
    pub fn close(&self, doc_id: &str) {
        let doc = self.docs.lock().unwrap().remove(doc_id);
        if let Some(doc) = doc {
            doc.generation.fetch_add(1, Ordering::SeqCst);
            doc.killer.kill();
            if let Ok(mut run) = doc.run.lock() {
                if let Some(mut w) = run.warm.take() {
                    let _ = w.run.child.kill();
                    let _ = w.run.child.wait();
                }
            }
            if let Ok(mut live) = doc.live.lock() {
                live.stop();
            }
        }
    }

    /// Every document's engines stop (the app quits: TeX processes do not outlive it).
    pub fn shutdown(&self) {
        let ids: Vec<String> = self.docs.lock().unwrap().keys().cloned().collect();
        for id in ids {
            self.close(&id);
        }
    }

    /// One paragraph typeset at once by the document's live server (live.rs), its glyphs and their outlines.
    pub fn live(&self, req: &crate::live::LiveRequest) -> crate::live::LiveResult {
        let doc = self.doc(&req.doc);
        let mut live = doc.live.lock().unwrap();
        live.paragraph(req)
    }

    pub fn compile(&self, req: &CompileRequest) -> CompileResult {
        let start = Instant::now();
        let Some(dist) = self.dist() else {
            return CompileResult::empty("notex");
        };
        let doc = self.doc(&req.doc);
        let my = doc.generation.fetch_add(1, Ordering::SeqCst) + 1;
        // an older compile of this document: its engine stops now
        doc.killer.kill();
        let mut run = doc.run.lock().unwrap();
        if doc.generation.load(Ordering::SeqCst) != my {
            return CompileResult::empty("superseded");
        }
        let mut result = CompileResult::empty("ok");
        if let Some(ticket) = &req.ticket {
            *doc.progress.lock().unwrap() = Progress { ticket: ticket.clone(), result: None, done: false };
        }
        let outcome = self.compile_locked(&dist, &doc, &mut run, req, my, &mut result);
        {
            let mut p = doc.progress.lock().unwrap();
            if req.ticket.as_deref() == Some(p.ticket.as_str()) {
                p.done = true;
            }
            doc.progress_cv.notify_all();
        }
        match outcome {
            Ok(()) => {}
            Err(e) => {
                if result.status == "ok" {
                    result.status = "failed".into();
                }
                result.error = e;
            }
        }
        result.timings.total = ms(start);
        result
    }

    fn superseded(doc: &Doc, my: u64) -> bool {
        doc.generation.load(Ordering::SeqCst) != my
    }

    #[allow(clippy::too_many_arguments)]
    fn compile_locked(&self, dist: &TexDist, doc: &Doc, run: &mut RunState, req: &CompileRequest, my: u64, result: &mut CompileResult) -> Result<(), String> {
        let root = self.allowed_root(&req.root)?;
        let main_rel = project::safe_rel(&req.main).ok_or("No main file")?;
        let work = doc.dir.join("work");
        let t = Instant::now();
        project::sync(&root, &work, &mut run.manifest, &req.overlays).map_err(|e| format!("copying the project: {e}"))?;
        result.timings.sync = ms(t);

        let main_path = work.join(&main_rel);
        let main_bytes = std::fs::read(&main_path).map_err(|e| format!("{}: {e}", req.main))?;
        // a main file that is not UTF-8 (GB18030) runs as it is: the warm engine's wrapper is written as UTF-8
        let utf8 = std::str::from_utf8(&main_bytes).is_ok();
        let main_text = String::from_utf8_lossy(&main_bytes).to_string();
        let files = project::list_files(&work);
        let (engine, reason) = match req.engine.as_deref().and_then(Engine::parse) {
            Some(e) => (e, "chosen".to_string()),
            None => engine::detect(&main_text, &self.class_texts(&work, &files)),
        };
        if !dist.has(engine.name()) {
            return Err(format!("{} is not installed", engine.name()));
        }
        result.engine = engine.name().into();
        result.engine_reason = reason;
        result.page_mode = if engine.output() == "pdf" || !dist.dvisvgm { "raster".into() } else { "svg".into() };
        // the main file's folder is where TeX runs (relative \input paths are the main file's)
        let run_dir = main_path.parent().unwrap_or(&work).to_path_buf();
        let job = main_rel.file_stem().map(|s| s.to_string_lossy().to_string()).unwrap_or("main".into());
        let split = if utf8 { engine::split_document(&main_text) } else { None };
        let key = self.preamble_key(engine, req.shell_escape, &main_text, split.as_ref(), &work, &files);

        let mut passes = 0u32;
        let mut warm_used = false;
        loop {
            passes += 1;
            let before = aux_hashes(&run_dir, &job);
            let t = Instant::now();
            // the first pass shows the visible pages as soon as TeX has written them (a long document's later pages
            // take a while yet)
            let early = if passes == 1 && req.ticket.is_some() && engine.output() != "pdf" && dist.dvisvgm && !req.visible.is_empty() {
                Some(Early { visible: req.visible.clone(), ticket: req.ticket.clone().unwrap_or_default(), start: Instant::now() })
            } else {
                None
            };
            let (status_ok, warm) = self.run_engine(dist, doc, run, engine, &run_dir, &job, &main_rel, split.as_ref(), key, req.shell_escape, my, early)?;
            if passes == 1 {
                warm_used = warm;
            }
            result.timings.tex += ms(t);
            if Self::superseded(doc, my) {
                result.status = "superseded".into();
                return Ok(());
            }
            let _ = status_ok;
            let after = aux_hashes(&run_dir, &job);
            let mut rerun = String::new();
            if before != after {
                rerun = "references".into();
            }
            // BibTeX / Biber
            let t = Instant::now();
            if let Some(tool) = self.bib_needed(run, dist, &run_dir, &job) {
                match self.run_bib(dist, &run_dir, &job, &tool, &doc.killer) {
                    Ok(()) => rerun = tool.clone(),
                    Err(e) => result.error = format!("{tool}: {e}"),
                }
                run.bib_key = bib_key(&run_dir, &job);
            }
            result.timings.bib += ms(t);
            run.aux = after;
            if !rerun.is_empty() && req.full && passes < 5 {
                continue;
            }
            if !rerun.is_empty() && !req.full {
                result.needs_rerun = true;
                result.rerun_reason = rerun;
            }
            break;
        }
        result.timings.passes = passes;
        result.warm = warm_used;
        let log_path = run_dir.join(format!("{job}.log"));
        result.log = String::from_utf8_lossy(&std::fs::read(&log_path).unwrap_or_default()).to_string();

        // the build: what this run produced
        let output = run_dir.join(format!("{job}.{}", engine.output()));
        if !output.exists() {
            result.status = "failed".into();
            self.spawn_warm(dist, run, engine, &run_dir, &job, split.as_ref(), key, req.shell_escape);
            return Ok(());
        }
        run.next_build += 1;
        let build_id = run.next_build;
        let bdir = doc.dir.join("builds").join(build_id.to_string());
        let _ = std::fs::remove_dir_all(&bdir);
        std::fs::create_dir_all(&bdir).map_err(|e| e.to_string())?;
        let out_name = format!("{job}.{}", engine.output());
        std::fs::copy(&output, bdir.join(&out_name)).map_err(|e| e.to_string())?;
        for ext in ["synctex.gz", "log", "fls"] {
            let _ = std::fs::copy(run_dir.join(format!("{job}.{ext}")), bdir.join(format!("{job}.{ext}")));
        }
        // the next edit's engine reads the preamble now
        self.spawn_warm(dist, run, engine, &run_dir, &job, split.as_ref(), key, req.shell_escape);
        // and the live server (the paragraph being edited, at once)
        doc.live.lock().unwrap().ensure(
            dist,
            &run_dir,
            &job,
            engine,
            split.as_ref(),
            key,
            req.shell_escape,
            &doc.dir.join("live"),
            &self.svg_font_format,
            default_paper(dist),
        );

        // the screen copy and the pages asked for
        let t = Instant::now();
        let mut pages = vec![];
        let mut pdf = None;
        if result.page_mode == "svg" {
            let bytes = std::fs::read(bdir.join(&out_name)).map_err(|e| e.to_string())?;
            let fls = std::fs::read_to_string(bdir.join(format!("{job}.fls"))).unwrap_or_default();
            let inputs = crate::fonts::recorded_inputs(&fls, &run_dir);
            let drawers = Drawers { mutool: dist.mutool.clone(), gs: dist.gs.clone(), extractbb: Some(dist.tool("extractbb")).filter(|p| p.is_file()) };
            let images = run.images.get_or_insert_with(|| ImageHost::new(&run_dir, &doc.dir.join("images"), drawers));
            images.dir = run_dir.clone();
            let kpse = dist.tool("kpsewhich");
            let paper = default_paper(dist);
            match screen::make_screen(&bytes, &bdir.join("screen.dvi"), &run_dir, &inputs, Some(&kpse), &mut run.tfm, images, paper) {
                Ok(p) => pages = p,
                Err(e) => {
                    result.error = format!("screen copy: {e}");
                    result.page_mode = "raster".into();
                }
            }
        }
        if result.page_mode == "raster" {
            let p = if engine.output() == "pdf" {
                bdir.join(&out_name)
            } else {
                self.make_pdf(dist, &run_dir, &bdir, &job, engine)?
            };
            pages = pdf_pages(&std::fs::read(&p).unwrap_or_default(), build_id);
            pdf = Some(p);
        }
        result.timings.screen = ms(t);
        let build = Arc::new(Build { dir: bdir, run_dir: run_dir.clone(), job, engine, pages: pages.clone(), pdf });
        {
            let mut view = doc.view.lock().unwrap();
            view.builds.insert(build_id, build.clone());
            while view.builds.len() > KEEP_BUILDS {
                let old = *view.builds.keys().next().unwrap();
                if let Some(b) = view.builds.remove(&old) {
                    let _ = std::fs::remove_dir_all(&b.dir);
                }
            }
        }
        result.build = build_id;
        result.pages = pages;
        if result.page_mode == "svg" && !req.visible.is_empty() {
            let t = Instant::now();
            let wanted: Vec<usize> = req.visible.iter().copied().filter(|p| *p >= 1 && *p <= build.pages.len()).collect();
            result.svgs = self.svg_pages(dist, doc, &build, &run_dir, &wanted);
            result.timings.svg = ms(t);
        }
        Ok(())
    }

    /// What decides whether the warm engine's preamble is still the document's.
    fn preamble_key(&self, engine: Engine, shell: bool, main: &str, split: Option<&engine::Split>, work: &Path, files: &[ProjectFile]) -> u64 {
        let mut h = Fnv::default();
        h.write(engine.name().as_bytes());
        h.write(&[shell as u8]);
        match split {
            Some(s) => h.write(s.preamble.as_bytes()),
            None => h.write(main.as_bytes()),
        }
        // the project's own classes, packages and the files the preamble reads
        let pre = split.map(|s| s.preamble.as_str()).unwrap_or("");
        for f in files {
            let l = f.path.to_ascii_lowercase();
            let stem = f.path.rsplit('/').next().unwrap_or("").rsplit_once('.').map(|(s, _)| s).unwrap_or("");
            let in_preamble = l.ends_with(".tex") && !stem.is_empty() && pre.contains(stem);
            if in_preamble || [".cls", ".sty", ".def", ".cfg", ".clo", ".fd", ".ldf", ".lua"].iter().any(|e| l.ends_with(e)) {
                h.write(f.path.as_bytes());
                h.write_u64(hash_bytes(&std::fs::read(work.join(&f.path)).unwrap_or_default()));
            }
        }
        h.finish()
    }

    #[allow(clippy::too_many_arguments)]
    fn spawn(&self, dist: &TexDist, engine: Engine, dir: &Path, job: &str, file: &str, shell: bool) -> std::io::Result<Running> {
        let mut cmd = Command::new(dist.tool(engine.name()));
        cmd.args(engine::arguments(engine, job, shell))
            .arg(file)
            .current_dir(dir)
            .env("PATH", dist.path_env())
            // long log lines: file:line:error messages stay on one line
            .env("max_print_line", "10000")
            .env("error_line", "254")
            .env("half_error_line", "238")
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::null());
        process::hide_window(&mut cmd);
        let mut child = cmd.spawn()?;
        let output = Arc::new(Mutex::new(Vec::new()));
        if let Some(mut out) = child.stdout.take() {
            let sink = output.clone();
            std::thread::spawn(move || {
                let mut buf = [0u8; 16384];
                while let Ok(n) = out.read(&mut buf) {
                    if n == 0 {
                        break;
                    }
                    let mut s = sink.lock().unwrap();
                    s.extend_from_slice(&buf[..n]);
                    if s.len() > 256 * 1024 {
                        let cut = s.len() - 128 * 1024;
                        s.drain(..cut);
                    }
                }
            });
        }
        let stdin = child.stdin.take();
        Ok(Running { child, stdin, output })
    }

    #[allow(clippy::too_many_arguments)]
    fn spawn_warm(&self, dist: &TexDist, run: &mut RunState, engine: Engine, dir: &Path, job: &str, split: Option<&engine::Split>, key: u64, shell: bool) {
        if let Some(mut w) = run.warm.take() {
            let _ = w.run.child.kill();
            let _ = w.run.child.wait();
        }
        let Some(split) = split else { return };
        let (wrapper, _) = engine::warm_files(engine, split);
        if std::fs::write(dir.join(engine::WARM_FILE), wrapper).is_err() {
            return;
        }
        if let Ok(r) = self.spawn(dist, engine, dir, job, engine::WARM_FILE, shell) {
            run.warm = Some(Warm { run: r, key });
        }
    }

    /// One engine pass. Returns (exit status ok, warm engine used).
    #[allow(clippy::too_many_arguments)]
    fn run_engine(
        &self,
        dist: &TexDist,
        doc: &Doc,
        run: &mut RunState,
        engine: Engine,
        dir: &Path,
        job: &str,
        main_rel: &Path,
        split: Option<&engine::Split>,
        key: u64,
        shell: bool,
        my: u64,
        mut early: Option<Early>,
    ) -> Result<(bool, bool), String> {
        let mut warm_used = false;
        let mut running = match split {
            Some(split) => {
                let (wrapper, body) = engine::warm_files(engine, split);
                std::fs::write(dir.join(engine::BODY_FILE), body).map_err(|e| e.to_string())?;
                let ready = match run.warm.take() {
                    Some(mut w) => {
                        if w.key == key && matches!(w.run.child.try_wait(), Ok(None)) {
                            warm_used = true;
                            Some(w.run)
                        } else {
                            let _ = w.run.child.kill();
                            let _ = w.run.child.wait();
                            None
                        }
                    }
                    None => None,
                };
                match ready {
                    Some(r) => r,
                    None => {
                        std::fs::write(dir.join(engine::WARM_FILE), wrapper).map_err(|e| e.to_string())?;
                        self.spawn(dist, engine, dir, job, engine::WARM_FILE, shell).map_err(|e| format!("{}: {e}", engine.name()))?
                    }
                }
            }
            None => {
                let file = main_rel.file_name().map(|s| s.to_string_lossy().to_string()).unwrap_or_default();
                self.spawn(dist, engine, dir, job, &file, shell).map_err(|e| format!("{}: {e}", engine.name()))?
            }
        };
        // the last run's output goes (it is kept in its build): what appears now is this run's
        let _ = std::fs::remove_file(dir.join(format!("{job}.{}", engine.output())));
        // the body: one terminal line (TeX reads it once the preamble is done); then end of input
        if let Some(mut stdin) = running.stdin.take() {
            if split.is_some() {
                let _ = stdin.write_all(engine::BODY_LINE.as_bytes());
                let _ = stdin.flush();
            }
            drop(stdin);
        }
        *doc.killer.0.lock().unwrap() = Some(running.child.id());
        let start = Instant::now();
        let output = dir.join(format!("{job}.{}", engine.output()));
        let mut seen_len = 0u64;
        let mut last_look = Instant::now();
        let status = loop {
            match running.child.try_wait() {
                Ok(Some(s)) => break Some(s),
                Ok(None) => {}
                Err(_) => break None,
            }
            if early.is_some() && last_look.elapsed() > Duration::from_millis(40) {
                last_look = Instant::now();
                let len = std::fs::metadata(&output).map(|m| m.len()).unwrap_or(0);
                if len != seen_len && len > 0 {
                    seen_len = len;
                    if self.try_early(dist, doc, run, dir, job, &output, early.as_ref().unwrap()) {
                        early = None;
                    }
                }
            }
            if Self::superseded(doc, my) || start.elapsed() > TEX_LIMIT {
                process::kill_pid(running.child.id());
                let _ = running.child.kill();
                let _ = running.child.wait();
                break None;
            }
            std::thread::sleep(Duration::from_millis(3));
        };
        *doc.killer.0.lock().unwrap() = None;
        let _ = running.output;
        Ok((status.map(|s| s.success()).unwrap_or(false), warm_used))
    }

    fn bib_needed(&self, run: &RunState, dist: &TexDist, dir: &Path, job: &str) -> Option<String> {
        let bcf = dir.join(format!("{job}.bcf"));
        let aux = std::fs::read_to_string(dir.join(format!("{job}.aux"))).unwrap_or_default();
        let tool = if bcf.exists() {
            "biber"
        } else if aux.contains("\\bibdata") {
            "bibtex"
        } else {
            return None;
        };
        if !dist.has(tool) {
            return None;
        }
        let bbl = dir.join(format!("{job}.bbl"));
        let key = bib_key(dir, job);
        if !bbl.exists() || key != run.bib_key {
            Some(tool.into())
        } else {
            None
        }
    }

    fn run_bib(&self, dist: &TexDist, dir: &Path, job: &str, tool: &str, killer: &Killer) -> Result<(), String> {
        let mut cmd = Command::new(dist.tool(tool));
        cmd.arg(job).current_dir(dir).env("PATH", dist.path_env());
        match process::run(cmd, Duration::from_secs(120), Some(killer)) {
            Ok((out, false)) if out.status.success() || dir.join(format!("{job}.bbl")).exists() => Ok(()),
            Ok((out, timed_out)) => {
                let text = format!("{}{}", String::from_utf8_lossy(&out.stdout), String::from_utf8_lossy(&out.stderr));
                let last = text.lines().rev().find(|l| !l.trim().is_empty()).unwrap_or(if timed_out { "timed out" } else { "failed" });
                Err(last.trim().chars().take(200).collect())
            }
            Err(e) => Err(e.to_string()),
        }
    }

    /// The PDF of a build: (x)dvipdfmx on its DVI/XDV (made once).
    fn make_pdf(&self, dist: &TexDist, run_dir: &Path, bdir: &Path, job: &str, engine: Engine) -> Result<PathBuf, String> {
        let pdf = bdir.join(format!("{job}.pdf"));
        if pdf.exists() {
            return Ok(pdf);
        }
        let input = bdir.join(format!("{job}.{}", engine.output()));
        let tool = if engine == Engine::Xelatex { "xdvipdfmx" } else { "dvipdfmx" };
        let mut cmd = Command::new(dist.tool(tool));
        cmd.arg("-q").arg("-o").arg(&pdf).arg(&input).current_dir(run_dir).env("PATH", dist.path_env());
        let (out, _) = process::run(cmd, Duration::from_secs(180), None).map_err(|e| format!("{tool}: {e}"))?;
        if !pdf.exists() {
            return Err(format!("{tool}: {}", String::from_utf8_lossy(&out.stderr).lines().last().unwrap_or("no PDF")));
        }
        Ok(pdf)
    }

    fn svg_pages(&self, dist: &TexDist, doc: &Doc, build: &Build, run_dir: &Path, pages: &[usize]) -> Vec<SvgPage> {
        let mut out = vec![];
        let mut missing = vec![];
        {
            let view = doc.view.lock().unwrap();
            for &p in pages {
                let meta = &build.pages[p - 1];
                match view.svg.get(&meta.hash_value) {
                    Some(svg) => out.push(SvgPage { page: p, hash: meta.hash.clone(), svg: svg.to_string() }),
                    None => missing.push((p, meta.hash_value)),
                }
            }
        }
        if !missing.is_empty() {
            let opts = SvgOptions {
                dvisvgm: dist.tool("dvisvgm"),
                font_format: self.svg_font_format.clone(),
                processes: std::thread::available_parallelism().map(|n| n.get()).unwrap_or(4).clamp(1, 6),
            };
            let made = svg::convert(&opts, run_dir, &build.dir.join("screen.dvi"), &missing, &build.dir.join("svg"));
            let mut view = doc.view.lock().unwrap();
            for (p, hash) in missing {
                if let Some(svg) = made.get(&p) {
                    let svg = Arc::new(svg.clone());
                    view.svg.insert(hash, svg.clone());
                    view.svg_order.push(hash);
                    out.push(SvgPage { page: p, hash: build.pages[p - 1].hash.clone(), svg: svg.to_string() });
                }
            }
            while view.svg_order.len() > SVG_CACHE {
                let old = view.svg_order.remove(0);
                if !view.svg_order.contains(&old) {
                    view.svg.remove(&old);
                }
            }
        }
        out.sort_by_key(|p| p.page);
        out
    }

    fn build(&self, doc_id: &str, build: u64) -> Result<(Arc<Doc>, Arc<Build>), String> {
        let doc = self.doc(doc_id);
        let b = doc.view.lock().unwrap().builds.get(&build).cloned().ok_or("That build is gone")?;
        Ok((doc, b))
    }

    /// Pages of a build: SVG, or a PNG of the PDF page when the SVG cannot show it exactly.
    pub fn pages(&self, req: &PagesRequest) -> Result<Vec<PageOut>, String> {
        let dist = self.dist().ok_or("TeX is not installed")?;
        let (doc, build) = self.build(&req.doc, req.build)?;
        let run_dir = build.run_dir.clone();
        let wanted: Vec<usize> = req.pages.iter().copied().filter(|p| *p >= 1 && *p <= build.pages.len()).collect();
        let raster_ok = dist.mutool.is_some() || dist.gs.is_some();
        let (svg_wanted, raster_wanted): (Vec<usize>, Vec<usize>) = wanted.iter().partition(|p| {
            let meta = &build.pages[**p - 1];
            build.engine.output() != "pdf" && (meta.unsupported.is_empty() || !raster_ok)
        });
        let mut out: Vec<PageOut> = self
            .svg_pages(&dist, &doc, &build, &run_dir, &svg_wanted)
            .into_iter()
            .map(|p| PageOut { page: p.page, hash: p.hash, svg: Some(p.svg), png: None })
            .collect();
        if !raster_wanted.is_empty() {
            let pdf = match &build.pdf {
                Some(p) => p.clone(),
                None => self.make_pdf(&dist, &run_dir, &build.dir, &build.job, build.engine)?,
            };
            let scale = req.scale.unwrap_or(2.0).clamp(0.5, 6.0);
            for p in raster_wanted {
                if let Some(png) = rasterize(&dist, &pdf, p, scale, &build.dir) {
                    out.push(PageOut { page: p, hash: build.pages[p - 1].hash.clone(), svg: None, png: Some(BASE64.encode(png)) });
                }
            }
        }
        out.sort_by_key(|p| p.page);
        Ok(out)
    }

    /// The PDF of a build (made from its DVI/XDV by (x)dvipdfmx the first time): its path, copied to `dest`,
    /// or its bytes.
    pub fn pdf(&self, req: &BuildRequest) -> Result<PdfOut, String> {
        let dist = self.dist().ok_or("TeX is not installed")?;
        let (_doc, build) = self.build(&req.doc, req.build)?;
        let run_dir = build.run_dir.clone();
        let pdf = match &build.pdf {
            Some(p) => p.clone(),
            None => self.make_pdf(&dist, &run_dir, &build.dir, &build.job, build.engine)?,
        };
        let size = std::fs::metadata(&pdf).map(|m| m.len()).unwrap_or(0);
        if let Some(dest) = &req.dest {
            let dest = PathBuf::from(dest);
            if !dest.is_absolute() {
                return Err("The export path must be absolute".into());
            }
            std::fs::copy(&pdf, &dest).map_err(|e| format!("{}: {e}", dest.display()))?;
        }
        let base64 = if req.bytes { Some(BASE64.encode(std::fs::read(&pdf).map_err(|e| e.to_string())?)) } else { None };
        Ok(PdfOut { path: pdf.to_string_lossy().to_string(), size, base64 })
    }

    pub fn synctex(&self, req: &BuildRequest) -> Result<SynctexOut, String> {
        let (_doc, build) = self.build(&req.doc, req.build)?;
        let bytes = std::fs::read(build.dir.join(format!("{}.synctex.gz", build.job))).map_err(|_| "No SyncTeX for this build".to_string())?;
        let work = build.run_dir.canonicalize().unwrap_or(build.run_dir.clone());
        Ok(SynctexOut {
            base64: BASE64.encode(bytes),
            work: work.to_string_lossy().to_string(),
            aliases: vec![engine::WARM_FILE.into(), engine::BODY_FILE.into()],
            main: String::new(),
        })
    }
}

/// What the first pass of a compile shows early.
struct Early {
    visible: Vec<usize>,
    ticket: String,
    start: Instant,
}

impl Service {
    /// The visible pages of the output TeX is still writing, once they are all there: converted and handed to
    /// `preview`. True when done (or when it cannot be).
    #[allow(clippy::too_many_arguments)]
    fn try_early(&self, dist: &TexDist, doc: &Doc, run: &mut RunState, dir: &Path, job: &str, output: &Path, early: &Early) -> bool {
        let Ok(bytes) = std::fs::read(output) else { return false };
        let Ok(mut dvi) = crate::dvi::Dvi::parse_partial(&bytes) else { return false };
        let need = early.visible.iter().copied().max().unwrap_or(1);
        // the page after the last visible one has begun: the visible ones are final
        if dvi.pages.len() < need {
            return false;
        }
        dvi.pages.truncate(need);
        let bodies: Vec<Vec<u8>> = dvi.pages.iter().map(|p| p.body.clone()).collect();
        let edir = doc.dir.join("early");
        let _ = std::fs::create_dir_all(&edir);
        let partial = edir.join("partial.dvi");
        if std::fs::write(&partial, dvi.write_with(&bodies)).is_err() {
            return true;
        }
        let fls = std::fs::read_to_string(dir.join(format!("{job}.fls"))).unwrap_or_default();
        let inputs = crate::fonts::recorded_inputs(&fls, dir);
        let drawers = Drawers { mutool: dist.mutool.clone(), gs: dist.gs.clone(), extractbb: Some(dist.tool("extractbb")).filter(|p| p.is_file()) };
        let images = run.images.get_or_insert_with(|| ImageHost::new(dir, &doc.dir.join("images"), drawers));
        let kpse = dist.tool("kpsewhich");
        let Ok(pages) = screen::make_screen(&std::fs::read(&partial).unwrap_or_default(), &edir.join("screen.dvi"), dir, &inputs, Some(&kpse), &mut run.tfm, images, default_paper(dist)) else {
            return true;
        };
        let build = Build { dir: edir.clone(), run_dir: dir.to_path_buf(), job: job.to_string(), engine: Engine::Xelatex, pages: pages.clone(), pdf: None };
        let svgs = self.svg_pages(dist, doc, &build, dir, &early.visible.iter().copied().filter(|p| *p >= 1 && *p <= pages.len()).collect::<Vec<_>>());
        let mut p = doc.progress.lock().unwrap();
        if p.ticket == early.ticket {
            p.result = Some(PreviewResult { status: "early".into(), pages, svgs, at: early.start.elapsed().as_millis() as u64 });
            doc.progress_cv.notify_all();
        }
        true
    }

    /// The early pages of compile `ticket` (waits until they are there, or the compile is over).
    pub fn preview(&self, req: &PreviewRequest) -> PreviewResult {
        let doc = self.doc(&req.doc);
        let mut p = doc.progress.lock().unwrap();
        let deadline = Instant::now() + Duration::from_secs(120);
        // the compile may not have begun yet (both requests are sent together): a while for it to come
        let arrival = Instant::now() + Duration::from_secs(10);
        loop {
            if p.ticket == req.ticket {
                if let Some(r) = &p.result {
                    return r.clone();
                }
                if p.done {
                    break;
                }
            } else if Instant::now() > arrival {
                break;
            }
            let left = deadline.saturating_duration_since(Instant::now());
            if left.is_zero() {
                break;
            }
            p = doc.progress_cv.wait_timeout(p, left.min(Duration::from_millis(500))).unwrap().0;
        }
        PreviewResult { status: "done".into(), ..Default::default() }
    }
}

/// Hashes of the files whose change means another pass (aux, contents lists, outlines, beamer navigation).
fn aux_hashes(dir: &Path, job: &str) -> HashMap<String, u64> {
    let mut out = HashMap::new();
    let exts = ["aux", "toc", "lof", "lot", "out", "nav", "snm", "loa", "thm", "bbl", "cpt", "ind"];
    let mut add = |p: PathBuf| {
        if let Ok(b) = std::fs::read(&p) {
            out.insert(p.to_string_lossy().to_string(), hash_bytes(&b));
        }
    };
    for e in exts {
        add(dir.join(format!("{job}.{e}")));
    }
    // \include-d chapters' aux files
    fn walk(d: &Path, depth: usize, add: &mut dyn FnMut(PathBuf)) {
        if depth > 3 {
            return;
        }
        if let Ok(rd) = std::fs::read_dir(d) {
            for e in rd.flatten() {
                let p = e.path();
                if p.is_dir() {
                    walk(&p, depth + 1, add);
                } else if p.extension().map(|x| x == "aux").unwrap_or(false)
                    // (the live server's own, which it rewrites on its own)
                    && !p.file_name().map(|n| n.to_string_lossy().starts_with("_vtx_")).unwrap_or(false)
                {
                    add(p);
                }
            }
        }
    }
    walk(dir, 0, &mut add);
    out
}

/// What BibTeX/Biber depend on: the citations and data named in .aux/.bcf and the .bib files.
fn bib_key(dir: &Path, job: &str) -> u64 {
    let mut h = Fnv::default();
    let aux = std::fs::read_to_string(dir.join(format!("{job}.aux"))).unwrap_or_default();
    for line in aux.lines() {
        if line.starts_with("\\citation") || line.starts_with("\\bibdata") || line.starts_with("\\bibstyle") || line.starts_with("\\abx@aux@cite") {
            h.write(line.as_bytes());
        }
    }
    if let Ok(b) = std::fs::read(dir.join(format!("{job}.bcf"))) {
        h.write(&b);
    }
    if let Ok(rd) = std::fs::read_dir(dir) {
        let mut bibs: Vec<PathBuf> = rd.flatten().map(|e| e.path()).filter(|p| p.extension().map(|x| x == "bib").unwrap_or(false)).collect();
        bibs.sort();
        for b in bibs {
            h.write_u64(hash_bytes(&std::fs::read(b).unwrap_or_default()));
        }
    }
    h.finish()
}

/// dvipdfmx's default paper (pages without a size special), looked up once.
fn default_paper(dist: &TexDist) -> (f64, f64) {
    static PAPER: std::sync::OnceLock<(f64, f64)> = std::sync::OnceLock::new();
    *PAPER.get_or_init(|| read_default_paper(dist))
}

fn read_default_paper(dist: &TexDist) -> (f64, f64) {
    let mut cmd = Command::new(dist.tool("kpsewhich"));
    cmd.arg("dvipdfmx.cfg");
    process::hide_window(&mut cmd);
    let path = cmd.output().ok().map(|o| String::from_utf8_lossy(&o.stdout).trim().to_string()).unwrap_or_default();
    let text = std::fs::read_to_string(path).unwrap_or_default();
    for line in text.lines() {
        if let Some(p) = line.trim().strip_prefix("p ") {
            return match p.trim() {
                "letter" => (612.0, 792.0),
                "a4" => (595.276, 841.89),
                _ => (595.276, 841.89),
            };
        }
    }
    (612.0, 792.0)
}

/// Page sizes of a PDF written without object compression (MediaBox per page, else inherited); one entry per
/// /Type /Page.
fn pdf_pages(pdf: &[u8], build: u64) -> Vec<PageMeta> {
    // (bytes searched as text: a PDF's structure is ASCII, its streams are skipped by the window sizes)
    let window = |from: usize, to: usize| String::from_utf8_lossy(&pdf[from.min(pdf.len())..to.min(pdf.len())]).to_string();
    let find = |hay: &[u8], needle: &[u8], from: usize| hay[from.min(hay.len())..].windows(needle.len()).position(|w| w == needle).map(|i| i + from);
    let mut default = (595.276, 841.89);
    if let Some(i) = find(pdf, b"/Type /Pages", 0).or_else(|| find(pdf, b"/Type/Pages", 0)) {
        if let Some(b) = find_box(&window(i.saturating_sub(400), i + 600)) {
            default = b;
        }
    }
    let mut out = vec![];
    let mut at = 0;
    let mut n = 0;
    while let Some(i) = find(pdf, b"/Type", at) {
        let after = window(i + 5, i + 16);
        let after = after.trim_start();
        if after.starts_with("/Page") && !after.starts_with("/Pages") {
            n += 1;
            // the page's own object: from its "obj" to its "endobj"
            let start = pdf[..i].windows(3).rposition(|w| w == b"obj").unwrap_or(i.saturating_sub(300));
            let end = find(pdf, b"endobj", i).unwrap_or(i + 400);
            let (w, h) = find_box(&window(start, end)).unwrap_or(default);
            out.push(PageMeta { width: w, height: h, hash: format!("{build}-{n}"), hash_value: build * 100000 + n, unsupported: vec![] });
        }
        at = i + 5;
    }
    out
}

fn find_box(s: &str) -> Option<(f64, f64)> {
    let i = s.find("/MediaBox")?;
    let rest = &s[i..];
    let (a, b) = (rest.find('[')?, rest.find(']')?);
    let v: Vec<f64> = rest[a + 1..b].split_whitespace().filter_map(|x| x.parse().ok()).collect();
    if v.len() == 4 {
        Some((v[2] - v[0], v[3] - v[1]))
    } else {
        None
    }
}

/// A PDF page as PNG at `scale` pixels per bp.
fn rasterize(dist: &TexDist, pdf: &Path, page: usize, scale: f64, dir: &Path) -> Option<Vec<u8>> {
    let out = dir.join(format!("raster-{page}-{}.png", (scale * 100.0).round() as u32));
    if let Ok(b) = std::fs::read(&out) {
        return Some(b);
    }
    let dpi = (72.0 * scale).round() as u32;
    if let Some(mutool) = &dist.mutool {
        let mut cmd = Command::new(mutool);
        cmd.args(["draw", "-q", "-r", &dpi.to_string(), "-o"]).arg(&out).arg(pdf).arg(page.to_string());
        let _ = process::run(cmd, Duration::from_secs(60), None);
    } else if let Some(gs) = &dist.gs {
        let mut cmd = Command::new(gs);
        cmd.args(["-q", "-dSAFER", "-dBATCH", "-dNOPAUSE", "-sDEVICE=png16m", "-dTextAlphaBits=4", "-dGraphicsAlphaBits=4"])
            .arg(format!("-r{dpi}"))
            .arg(format!("-dFirstPage={page}"))
            .arg(format!("-dLastPage={page}"))
            .arg(format!("-sOutputFile={}", out.display()))
            .arg(pdf)
            .env("PATH", dist.path_env());
        let _ = process::run(cmd, Duration::from_secs(60), None);
    }
    std::fs::read(&out).ok()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn page_boxes_from_uncompressed_pdf() {
        let pdf = b"%PDF-1.5\n1 0 obj << /Type /Pages /Kids [2 0 R 3 0 R] /MediaBox [0 0 612 792] >> endobj\n2 0 obj << /Type /Page /Parent 1 0 R >> endobj\n3 0 obj << /Type /Page /MediaBox [0 0 300 400] /Parent 1 0 R >> endobj";
        let pages = pdf_pages(pdf, 1);
        assert_eq!(pages.len(), 2);
        assert_eq!((pages[0].width, pages[1].height), (612.0, 400.0));
    }
}
