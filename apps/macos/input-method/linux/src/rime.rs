//! librime with VisualTeX's pinyin schema, as the macOS input method uses it
//! (input-method/Sources/RimeEngine.swift) and the Windows one (server/src/rime.rs there, which this
//! follows). Its shared data (schema, dictionaries, prebuilt tables) is in $VISUALTEX_RIME_DATA, else
//! share/visualtex/rime beside this program's prefix (/usr/share/visualtex/rime for the package);
//! what it learns (the user dictionary) goes to $XDG_DATA_HOME/visualtex/rime.

use std::ffi::{c_char, c_int, c_void, CStr, CString};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::OnceLock;

pub type SessionId = usize;
type Bool = c_int;

#[repr(C)]
struct RimeTraits {
    data_size: c_int,
    shared_data_dir: *const c_char,
    user_data_dir: *const c_char,
    distribution_name: *const c_char,
    distribution_code_name: *const c_char,
    distribution_version: *const c_char,
    app_name: *const c_char,
    modules: *const *const c_char,
    min_log_level: c_int,
    log_dir: *const c_char,
    prebuilt_data_dir: *const c_char,
    staging_dir: *const c_char,
}

#[repr(C)]
struct RimeComposition {
    length: c_int,
    cursor_pos: c_int,
    sel_start: c_int,
    sel_end: c_int,
    preedit: *mut c_char,
}

#[repr(C)]
struct RimeCandidate {
    text: *mut c_char,
    comment: *mut c_char,
    reserved: *mut c_void,
}

#[repr(C)]
struct RimeMenu {
    page_size: c_int,
    page_no: c_int,
    is_last_page: Bool,
    highlighted_candidate_index: c_int,
    num_candidates: c_int,
    candidates: *mut RimeCandidate,
    select_keys: *mut c_char,
}

#[repr(C)]
struct RimeCommit {
    data_size: c_int,
    text: *mut c_char,
}

#[repr(C)]
struct RimeContext {
    data_size: c_int,
    composition: RimeComposition,
    menu: RimeMenu,
    commit_text_preview: *mut c_char,
    select_labels: *mut *mut c_char,
}

type Unused = Option<unsafe extern "C" fn()>;

/// The first members of librime's RimeApi (rime_api.h), in its order: only a prefix is read.
#[repr(C)]
struct RimeApi {
    data_size: c_int,
    setup: Option<unsafe extern "C" fn(*mut RimeTraits)>,
    set_notification_handler: Option<unsafe extern "C" fn(*const c_void, *mut c_void)>,
    initialize: Option<unsafe extern "C" fn(*mut RimeTraits)>,
    finalize: Option<unsafe extern "C" fn()>,
    start_maintenance: Option<unsafe extern "C" fn(Bool) -> Bool>,
    is_maintenance_mode: Option<unsafe extern "C" fn() -> Bool>,
    join_maintenance_thread: Option<unsafe extern "C" fn()>,
    deployer_initialize: Option<unsafe extern "C" fn(*mut RimeTraits)>,
    prebuild: Option<unsafe extern "C" fn() -> Bool>,
    deploy: Option<unsafe extern "C" fn() -> Bool>,
    deploy_schema: Option<unsafe extern "C" fn(*const c_char) -> Bool>,
    deploy_config_file: Option<unsafe extern "C" fn(*const c_char, *const c_char) -> Bool>,
    sync_user_data: Option<unsafe extern "C" fn() -> Bool>,
    create_session: Option<unsafe extern "C" fn() -> SessionId>,
    find_session: Option<unsafe extern "C" fn(SessionId) -> Bool>,
    destroy_session: Option<unsafe extern "C" fn(SessionId) -> Bool>,
    cleanup_stale_sessions: Option<unsafe extern "C" fn()>,
    cleanup_all_sessions: Option<unsafe extern "C" fn()>,
    process_key: Option<unsafe extern "C" fn(SessionId, c_int, c_int) -> Bool>,
    commit_composition: Option<unsafe extern "C" fn(SessionId) -> Bool>,
    clear_composition: Option<unsafe extern "C" fn(SessionId)>,
    get_commit: Option<unsafe extern "C" fn(SessionId, *mut RimeCommit) -> Bool>,
    free_commit: Option<unsafe extern "C" fn(*mut RimeCommit) -> Bool>,
    get_context: Option<unsafe extern "C" fn(SessionId, *mut RimeContext) -> Bool>,
    free_context: Option<unsafe extern "C" fn(*mut RimeContext) -> Bool>,
    get_status: Option<unsafe extern "C" fn(SessionId, *mut c_void) -> Bool>,
    free_status: Option<unsafe extern "C" fn(*mut c_void) -> Bool>,
    set_option: Option<unsafe extern "C" fn(SessionId, *const c_char, Bool)>,
    get_option: Option<unsafe extern "C" fn(SessionId, *const c_char) -> Bool>,
    // set_property … set_caret_pos (30–74 in rime_api.h)
    unused: [Unused; 45],
    select_candidate_on_current_page: Option<unsafe extern "C" fn(SessionId, usize) -> Bool>,
}

#[link(name = "rime")]
extern "C" {
    fn rime_get_api() -> *const RimeApi;
}

fn data_size<T>() -> c_int {
    (std::mem::size_of::<T>() - std::mem::size_of::<c_int>()) as c_int
}

/// What one session shows: the spelling being composed, its caret (a character index) and the
/// candidate page.
#[derive(Clone, Debug, Default)]
pub struct State {
    pub preedit: String,
    pub cursor: usize,
    pub candidates: Vec<String>,
    pub comments: Vec<String>,
    pub labels: Vec<String>,
    pub highlighted: usize,
}

impl State {
    pub fn is_composing(&self) -> bool {
        !self.preedit.is_empty()
    }
}

pub struct Engine {
    api: &'static RimeApi,
}

/// The prebuilt tables are installed (librime's maintenance finished).
static READY: AtomicBool = AtomicBool::new(false);

unsafe impl Send for Engine {}
unsafe impl Sync for Engine {}

static ENGINE: OnceLock<Option<Engine>> = OnceLock::new();

fn utf8_path(path: &Path) -> CString {
    CString::new(path.to_string_lossy().into_owned()).unwrap_or_default()
}

fn text(pointer: *const c_char) -> String {
    if pointer.is_null() {
        return String::new();
    }
    unsafe { CStr::from_ptr(pointer) }.to_string_lossy().into_owned()
}

/// RIME's shared data: $VISUALTEX_RIME_DATA, else share/visualtex/rime in this program's prefix
/// (the program is in lib/visualtex: /usr/lib/visualtex/visualtex-ibus → /usr/share/visualtex/rime).
pub fn shared_data_dir() -> PathBuf {
    if let Some(directory) = std::env::var_os("VISUALTEX_RIME_DATA") {
        return PathBuf::from(directory);
    }
    let program = std::env::current_exe().unwrap_or_default();
    let prefix = program.parent().and_then(Path::parent).and_then(Path::parent).unwrap_or(Path::new("/usr"));
    prefix.join("share/visualtex/rime")
}

pub fn user_data_dir() -> PathBuf {
    let base = std::env::var_os("XDG_DATA_HOME")
        .map(PathBuf::from)
        .or_else(|| std::env::var_os("HOME").map(|home| PathBuf::from(home).join(".local/share")))
        .unwrap_or_else(std::env::temp_dir);
    base.join("visualtex").join("rime")
}

/// The engine, set up on first use (None when librime has no API).
pub fn engine() -> Option<&'static Engine> {
    ENGINE.get_or_init(start).as_ref()
}

fn start() -> Option<Engine> {
    let api = unsafe { rime_get_api().as_ref() }?;
    let shared = shared_data_dir();
    let user = user_data_dir();
    let _ = std::fs::create_dir_all(&user);
    // librime keeps these pointers: they live as long as the process.
    let keep = |value: CString| -> *const c_char { Box::leak(value.into_boxed_c_str()).as_ptr() };
    let mut traits = RimeTraits {
        data_size: data_size::<RimeTraits>(),
        shared_data_dir: keep(utf8_path(&shared)),
        user_data_dir: keep(utf8_path(&user)),
        distribution_name: keep(CString::new("VisualTeX").unwrap()),
        distribution_code_name: keep(CString::new("visualtex").unwrap()),
        distribution_version: keep(CString::new(env!("CARGO_PKG_VERSION")).unwrap()),
        app_name: keep(CString::new("rime.visualtex").unwrap()),
        modules: std::ptr::null(),
        min_log_level: 2,
        log_dir: keep(CString::new("").unwrap()),
        prebuilt_data_dir: keep(utf8_path(&shared.join("build"))),
        staging_dir: keep(utf8_path(&user.join("build"))),
    };
    unsafe {
        (api.setup?)(&mut traits);
        (api.initialize?)(&mut traits);
    }
    let engine = Engine { api };
    // The prebuilt tables are installed in the background: keys typed before then are not taken.
    let started = std::time::Instant::now();
    let maintenance = api.start_maintenance;
    let join = api.join_maintenance_thread;
    let shared_text = shared.display().to_string();
    std::thread::spawn(move || {
        unsafe {
            if let (Some(start), Some(join)) = (maintenance, join) {
                if start(0) != 0 {
                    join();
                }
            }
        }
        READY.store(true, Ordering::SeqCst);
        crate::log::write(&format!("rime ready in {} ms (data {shared_text})", started.elapsed().as_millis()));
    });
    Some(engine)
}

impl Engine {
    pub fn is_ready(&self) -> bool {
        READY.load(Ordering::SeqCst)
    }

    pub fn create_session(&self) -> SessionId {
        if !self.is_ready() {
            return 0;
        }
        self.api.create_session.map(|create| unsafe { create() }).unwrap_or(0)
    }

    pub fn destroy_session(&self, session: SessionId) {
        if session != 0 {
            if let Some(destroy) = self.api.destroy_session {
                unsafe { destroy(session) };
            }
        }
    }

    /// True when librime took the key (an X keysym and modifier mask, as IBus gives them).
    pub fn process(&self, session: SessionId, keysym: u32, mask: u32) -> bool {
        session != 0
            && self
                .api
                .process_key
                .map(|process| unsafe { process(session, keysym as c_int, mask as c_int) } != 0)
                .unwrap_or(false)
    }

    /// The text librime committed with the last key, if any.
    pub fn take_commit(&self, session: SessionId) -> Option<String> {
        let (get, free) = (self.api.get_commit?, self.api.free_commit?);
        let mut commit = RimeCommit { data_size: data_size::<RimeCommit>(), text: std::ptr::null_mut() };
        if unsafe { get(session, &mut commit) } == 0 {
            return None;
        }
        let result = text(commit.text);
        unsafe { free(&mut commit) };
        Some(result)
    }

    pub fn state(&self, session: SessionId) -> State {
        let mut state = State::default();
        let (Some(get), Some(free)) = (self.api.get_context, self.api.free_context) else { return state };
        if session == 0 {
            return state;
        }
        let mut context: RimeContext = unsafe { std::mem::zeroed() };
        context.data_size = data_size::<RimeContext>();
        if unsafe { get(session, &mut context) } == 0 {
            return state;
        }
        let composition = &context.composition;
        if !composition.preedit.is_null() && composition.length > 0 {
            let bytes = unsafe { std::slice::from_raw_parts(composition.preedit as *const u8, composition.length as usize) };
            state.preedit = String::from_utf8_lossy(bytes).into_owned();
            let offset = (composition.cursor_pos.max(0) as usize).min(bytes.len());
            state.cursor = String::from_utf8_lossy(&bytes[..offset]).chars().count();
        }
        let menu = &context.menu;
        if !menu.candidates.is_null() && menu.num_candidates > 0 {
            let candidates = unsafe { std::slice::from_raw_parts(menu.candidates, menu.num_candidates as usize) };
            state.candidates = candidates.iter().map(|candidate| text(candidate.text)).collect();
            state.comments = candidates.iter().map(|candidate| text(candidate.comment)).collect();
        }
        if !menu.select_keys.is_null() {
            state.labels = text(menu.select_keys).chars().map(String::from).collect();
        }
        if state.labels.len() < state.candidates.len() {
            state.labels = (1..=state.candidates.len()).map(|number| (number % 10).to_string()).collect();
        }
        state.highlighted = menu.highlighted_candidate_index.max(0) as usize;
        unsafe { free(&mut context) };
        state
    }

    pub fn is_ascii(&self, session: SessionId) -> bool {
        let Some(get) = self.api.get_option else { return false };
        session != 0 && unsafe { get(session, c"ascii_mode".as_ptr()) != 0 }
    }

    pub fn set_ascii(&self, session: SessionId, on: bool) {
        if let (Some(set), true) = (self.api.set_option, session != 0) {
            unsafe { set(session, c"ascii_mode".as_ptr(), on as Bool) };
        }
    }

    pub fn clear(&self, session: SessionId) {
        if let (Some(clear), true) = (self.api.clear_composition, session != 0) {
            unsafe { clear(session) };
        }
    }

    /// Commits what is being composed as it is spelled.
    pub fn commit_composition(&self, session: SessionId) -> Option<String> {
        let commit = self.api.commit_composition?;
        if session == 0 || unsafe { commit(session) } == 0 {
            return None;
        }
        self.take_commit(session)
    }

    /// A candidate of the page shown (a click in the candidate list).
    pub fn select_on_page(&self, session: SessionId, index: usize) -> bool {
        // the member is there only in a librime as new as rime_api.h's 75th function
        let offset = std::mem::offset_of!(RimeApi, select_candidate_on_current_page) - std::mem::size_of::<c_int>();
        if (self.api.data_size as usize) < offset + std::mem::size_of::<usize>() {
            return false;
        }
        session != 0
            && self
                .api
                .select_candidate_on_current_page
                .map(|select| unsafe { select(session, index) } != 0)
                .unwrap_or(false)
    }
}
