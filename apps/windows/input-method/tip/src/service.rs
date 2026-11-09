//! The text service: one per UI thread of an application that types with VisualTeX. VisualTeX is
//! a whole input method, as on the Mac: Chinese through pinyin (RIME, in VisualTeX Input) and
//! English (Shift switches), with formulas as one more mode. Its sinks share one state (`Shared`).
//!
//! Pinyin: the spelling is the field's composition (underlined), the candidates are VisualTeX
//! Input's window under it, and the chosen text replaces the composition.
//!
//! A formula (VisualTeX's "session on", from its shortcut): the field gets a composition where the
//! caret was holding the formula's space (space.rs: a selected hair space, which hides the field's
//! caret, then blanks as wide as each piece of the formula, joined by word spaces where the field
//! may break the line), so the field wraps the formula as its own text; the TIP reads back where
//! each piece went and VisualTeX draws it there. Every key but the system's goes to VisualTeX
//! (pinyin spelled in the formula's text goes to it as its composition), and "commit" replaces the
//! composition with the formula's text.
//!
//! An application typing through IMM32 (Qt's, such as WeChat) sees the composition through
//! Windows' translation (CUAS, a "transitory" context): the anchor there is a zero-width space
//! marked as the clause being converted (the field hides its caret for it and draws nothing), and
//! the field's caret is the one the application gave IMM32 for the candidate window, where the
//! context's own rectangles are made up. Chromium's own text store (Chrome, Edge, WebView2,
//! Electron) calls itself transitory too, but it lays the composition out as its text and tells
//! where each character went: the formula wraps there as in any other field.

use crate::engine::{Engine, Reply};
use crate::link::{Link, WM_LINK_CLOSED, WM_LINK_MESSAGE};
use crate::{display, edit, keys, log, profile, rimekeys, space};
use serde_json::{json, Value};
use std::cell::{Cell, RefCell};
use std::collections::HashMap;
use std::mem::ManuallyDrop;
use std::rc::Rc;
use windows::core::{implement, w, Interface, Ref, Result, BOOL, GUID};
use windows::Win32::Foundation::{E_INVALIDARG, FALSE, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::ClientToScreen;
use windows::Win32::System::Variant::{VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0, VT_I4};
use windows::Win32::UI::HiDpi::LogicalToPhysicalPointForPerMonitorDPI;
use windows::Win32::UI::Input::Ime::{ImmGetCandidateWindow, ImmGetContext, ImmReleaseContext, CANDIDATEFORM, CFS_EXCLUDE};
use windows::Win32::UI::Input::KeyboardAndMouse::GetFocus;
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};
use windows::Win32::UI::TextServices::{
    CLSID_TF_CategoryMgr, IEnumTfDisplayAttributeInfo, ITfCategoryMgr, ITfComposition, ITfCompositionSink,
    ITfCompositionSink_Impl, ITfContext, ITfContextComposition, ITfDisplayAttributeInfo, ITfDisplayAttributeProvider,
    ITfDisplayAttributeProvider_Impl, ITfDocumentMgr, ITfInsertAtSelection, ITfKeyEventSink, ITfKeyEventSink_Impl,
    ITfKeystrokeMgr, ITfRange, ITfSource, ITfTextInputProcessor, ITfTextInputProcessorEx,
    ITfTextInputProcessorEx_Impl, ITfTextInputProcessor_Impl, ITfThreadMgr, ITfThreadMgrEventSink,
    ITfThreadMgrEventSink_Impl, GUID_PROP_ATTRIBUTE, INSERT_TEXT_AT_SELECTION_FLAGS, TF_AE_NONE, TF_ANCHOR_END,
    TF_ANCHOR_START, TF_DEFAULT_SELECTION, TF_IAS_QUERYONLY, TF_SELECTION, TF_SELECTIONSTYLE, TF_SS_TRANSITORY,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetClassNameW, GetWindowLongPtrW, KillTimer,
    PostMessageW, RegisterClassW, SetTimer, SetWindowLongPtrW, GWLP_USERDATA, HWND_MESSAGE, WINDOW_EX_STYLE,
    WINDOW_STYLE, WM_APP, WM_TIMER, WNDCLASSW,
};

/// The formula's place in a field typing through IMM32: one zero-width space (the field draws the
/// clause being converted highlighted, Qt in its text colour: the anchor must have no width).
const IMM_ANCHOR: &[u16] = &[0x200B];
/// The formula's anchor is followed (Chrome lays the composition out a frame or so later; the
/// field may scroll): fast at first, then slowly while the formula is open.
const POLL_TIMER: usize = 1;
const FAST_POLLS: u32 = 30;
/// The candidates follow the pinyin composition for a moment after each key (a late layout).
const PLACE_TIMER: usize = 2;
const PLACE_POLLS: u32 = 8;
/// While VisualTeX is not linked, it is looked for every second (it may start later).
const LINK_TIMER: usize = 3;
/// Hands the field back to its previous input method, after the formula's last edit.
const WM_HAND_BACK: u32 = WM_APP + 0x53;

type Edges = (i32, i32, i32, i32);

#[derive(Default)]
struct Shared {
    thread_mgr: RefCell<Option<ITfThreadMgr>>,
    client: Cell<u32>,
    thread_sink_cookie: Cell<u32>,
    window: Cell<isize>,
    link: RefCell<Option<Link>>,
    key_sink: RefCell<Option<ITfKeyEventSink>>,
    composition_sink: RefCell<Option<ITfCompositionSink>>,
    engine: Engine,
    /// The display attribute of pinyin being spelled (0: none registered).
    display_atom: Cell<u32>,
    /// The display attribute of the formula's anchor in a field typing through IMM32 (0: none).
    anchor_atom: Cell<u32>,
    /// The display attribute of the space a field holds for a formula (0: none).
    space_atom: Cell<u32>,
    /// Chinese (false) or English, as RIME last said.
    ascii: Cell<bool>,
    /// A key RIME was offered in OnTestKey*, for the OnKey* that follows: (vk, release, reply).
    pending: RefCell<Option<(u32, bool, Reply)>>,

    // Pinyin in the field.
    pinyin: RefCell<Option<ITfComposition>>,
    pinyin_context: RefCell<Option<ITfContext>>,
    placed: Cell<Option<Edges>>,
    place_polls: Cell<u32>,

    // The formula.
    /// A formula is open in this thread's focused field.
    composing: Cell<bool>,
    formula: RefCell<Option<ITfComposition>>,
    formula_context: RefCell<Option<ITfContext>>,
    /// The formula's field types through IMM32 (a transitory context).
    formula_transitory: Cell<bool>,
    /// What the formula inserts, for a field that ends the composition itself.
    formula_text: RefCell<String>,
    /// The formula's caret from its start, physical pixels, and down to the caret's line.
    formula_caret_x: Cell<i32>,
    formula_caret_y: Cell<i32>,
    /// The formula's pieces (physical pixels wide, from VisualTeX), the composition holding them
    /// (space.rs), and the widths of a filler and of a joint as this field draws them.
    piece_widths: RefCell<Vec<f64>>,
    space: RefCell<space::Layout>,
    filler_width: Cell<f64>,
    joint_width: Cell<f64>,
    /// Calibrations in this formula: a few at most (a field rounding places to pixels must not make
    /// the space flip between two widths).
    calibrations: Cell<u32>,
    /// Where each piece started at the last poll, and what was last reported.
    last_pieces: RefCell<Vec<Edges>>,
    reported_pieces: RefCell<Option<Value>>,
    /// The widths changed at the last poll: the space is held again.
    rehold: Cell<bool>,
    /// Pinyin is being spelled in the formula's text (as its composition).
    formula_pinyin: Cell<bool>,
    /// Chinese or English before the formula opened (a formula starts in English).
    ascii_before_formula: Cell<Option<bool>>,
    last_rect: Cell<Option<Edges>>,
    polls: Cell<u32>,
    stable_rect: Cell<Option<Edges>>,
    /// The input method to hand the field back to after the formula.
    hand_back: RefCell<Option<Value>>,
    /// VisualTeX's screenshot selection has the foreground for a moment: the formula stays
    /// open when the field loses the keyboard meanwhile.
    hold: Cell<bool>,
}

impl Shared {
    fn window(&self) -> HWND {
        HWND(self.window.get() as *mut _)
    }

    /// Sends a message to VisualTeX (reconnecting when it restarted).
    fn send(&self, message: Value) -> bool {
        let sent = self.link.borrow().as_ref().map(|link| link.send(&message)).unwrap_or(false);
        if sent {
            return true;
        }
        if let Some(link) = Link::connect(self.window()) {
            link.send(&json!({ "t": "hello", "version": env!("CARGO_PKG_VERSION") }));
            let sent = link.send(&message);
            *self.link.borrow_mut() = Some(link);
            return sent;
        }
        false
    }

    fn focused_context(&self) -> Option<ITfContext> {
        let manager = self.thread_mgr.borrow().clone()?;
        let document: ITfDocumentMgr = unsafe { manager.GetFocus() }.ok()?;
        unsafe { document.GetTop() }.ok()
    }
}

fn client_name() -> String {
    std::env::current_exe()
        .ok()
        .and_then(|exe| exe.file_stem().map(|name| name.to_string_lossy().into_owned()))
        .unwrap_or_default()
}

fn utf16(text: &str) -> Vec<u16> {
    text.encode_utf16().collect()
}

/// Where a range of the field's text is on the screen.
struct Extent {
    /// Physical pixels (top-left origin), whatever the application's DPI awareness.
    rect: RECT,
    /// The field cut the rectangle to what it shows (a line partly scrolled out of view is shorter
    /// than its text).
    clipped: bool,
    /// The view's window (0: none).
    window: isize,
}

/// The screen rectangle of `range` (see Extent).
fn range_extent(context: &ITfContext, cookie: u32, range: &ITfRange) -> Option<Extent> {
    let view = unsafe { context.GetActiveView() }.ok()?;
    let mut rect = RECT::default();
    let mut clipped = BOOL(0);
    unsafe { view.GetTextExt(cookie, range, &mut rect, &mut clipped) }.ok()?;
    if rect.bottom <= rect.top {
        return None;
    }
    let window = unsafe { view.GetWnd() }.ok().filter(|window| !window.is_invalid());
    if let Some(window) = window {
        let mut start = POINT { x: rect.left, y: rect.top };
        let mut end = POINT { x: rect.right, y: rect.bottom };
        let converted = unsafe {
            LogicalToPhysicalPointForPerMonitorDPI(Some(window), &mut start).as_bool()
                && LogicalToPhysicalPointForPerMonitorDPI(Some(window), &mut end).as_bool()
        };
        // Both corners or neither: one converted and one not would make a wrong height.
        if converted {
            rect = RECT { left: start.x, top: start.y, right: end.x, bottom: end.y };
        }
    }
    Some(Extent { rect, clipped: clipped.as_bool(), window: window.map(|window| window.0 as isize).unwrap_or(0) })
}

/// The screen rectangle of `range` in physical pixels (top-left origin).
fn range_rect(context: &ITfContext, cookie: u32, range: &ITfRange) -> Option<RECT> {
    range_extent(context, cookie, range).map(|extent| extent.rect)
}

/// The context is the one Windows keeps for an application typing through IMM32 (CUAS): it holds
/// the composition only, and its rectangles are made up from the application's composition window
/// (one line of the default font below the caret where the application gives no font, as Qt).
fn is_transitory(context: &ITfContext) -> bool {
    unsafe { context.GetStatus() }.map(|status| status.dwStaticFlags & TF_SS_TRANSITORY != 0).unwrap_or(false)
}

/// The window of the context's view (0: none).
fn view_window(context: &ITfContext) -> isize {
    unsafe { context.GetActiveView().and_then(|view| view.GetWnd()) }
        .ok()
        .filter(|window| !window.is_invalid())
        .map(|window| window.0 as isize)
        .unwrap_or(0)
}

/// The view's window is Chromium's (Chrome_WidgetWin_1, Chrome_RenderWidgetHostHWND: Chrome, Edge,
/// WebView2, Electron applications).
fn chromium_view(context: &ITfContext) -> bool {
    let window = view_window(context);
    if window == 0 {
        return false;
    }
    let mut name = [0u16; 64];
    let length = unsafe { GetClassNameW(HWND(window as *mut _), &mut name) };
    let length = usize::try_from(length).unwrap_or(0).min(name.len());
    String::from_utf16_lossy(&name[..length]).starts_with("Chrome_")
}

/// The field types through IMM32 (see is_transitory). Chromium's text store calls itself
/// transitory as well (it keeps no text for reconversion), but its rectangles are its layout's,
/// one per character of the composition, so the formula's space is held there as in any field.
fn types_through_imm32(context: &ITfContext) -> bool {
    is_transitory(context) && !chromium_view(context)
}

/// The caret an application typing through IMM32 gave Windows for the candidate window: the area
/// the candidates must not cover (Qt sets its text cursor's rectangle there), in physical screen
/// pixels. Only a caret-like area counts (some applications exclude their whole text area). Read in
/// the application's own UI thread, as IMM32 requires.
fn imm_caret(window: isize) -> Option<RECT> {
    let focus = unsafe { GetFocus() };
    for window in [HWND(window as *mut _), focus] {
        if window.is_invalid() {
            continue;
        }
        let context = unsafe { ImmGetContext(window) };
        if context.0.is_null() {
            continue;
        }
        let mut form = CANDIDATEFORM::default();
        let read = unsafe { ImmGetCandidateWindow(context, 0, &mut form) }.as_bool();
        unsafe {
            let _ = ImmReleaseContext(window, context);
        }
        let area = form.rcArea;
        let (width, height) = (area.right - area.left, area.bottom - area.top);
        if !read || form.dwStyle != CFS_EXCLUDE || !(4..=400).contains(&height) || !(0..=height).contains(&width) {
            continue;
        }
        let mut start = POINT { x: area.left, y: area.top };
        let mut end = POINT { x: area.right, y: area.bottom };
        if !unsafe { ClientToScreen(window, &mut start).as_bool() && ClientToScreen(window, &mut end).as_bool() } {
            continue;
        }
        // Both corners or neither (see range_extent).
        let (mut physical_start, mut physical_end) = (start, end);
        let converted = unsafe {
            LogicalToPhysicalPointForPerMonitorDPI(Some(window), &mut physical_start).as_bool()
                && LogicalToPhysicalPointForPerMonitorDPI(Some(window), &mut physical_end).as_bool()
        };
        if converted {
            (start, end) = (physical_start, physical_end);
        }
        return Some(RECT { left: start.x, top: start.y, right: end.x, bottom: end.y });
    }
    None
}

/// A transitory context's rectangle is made up where, at the IMM32 caret's place, it is not the
/// caret's line: it lies mostly outside the caret (Qt's: below it) or is of another height. One far
/// from the caret is about another place, and stays.
fn made_up(rect: &RECT, caret: &RECT) -> bool {
    let caret_height = caret.bottom - caret.top;
    let height = rect.bottom - rect.top;
    let overlap = rect.bottom.min(caret.bottom) - rect.top.max(caret.top);
    let near = (rect.left - caret.left).abs() <= caret_height * 2 && (rect.top - caret.top).abs() <= caret_height * 2;
    near && (overlap * 2 < caret_height || height * 4 < caret_height * 3 || height * 3 > caret_height * 4)
}

/// The extent to report for a field: `extent` (the context's), or in a transitory context the
/// IMM32 caret where that is made up. Some(note): the caret is IMM32's (the note says what it
/// replaced, for the log).
fn field_extent(context: &ITfContext, extent: Option<Extent>) -> (Option<Extent>, Option<String>) {
    if !types_through_imm32(context) {
        return (extent, None);
    }
    let window = extent.as_ref().map(|extent| extent.window).filter(|&window| window != 0).unwrap_or_else(|| view_window(context));
    match (imm_caret(window), extent) {
        (Some(caret), Some(extent)) if !made_up(&extent.rect, &caret) => (Some(extent), None),
        (Some(caret), replaced) => {
            let note = match replaced {
                Some(Extent { rect, .. }) => {
                    format!("imm (made up {},{} {}x{})", rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)
                }
                None => "imm (no rectangle)".to_string(),
            };
            (Some(Extent { rect: caret, clipped: false, window }), Some(note))
        }
        (None, extent) => (extent, None),
    }
}

fn edges(rect: &RECT) -> Edges {
    (rect.left, rect.top, rect.right, rect.bottom)
}

fn selection_range(context: &ITfContext, cookie: u32) -> Option<ITfRange> {
    let mut selection = [TF_SELECTION::default()];
    let mut fetched = 0u32;
    unsafe { context.GetSelection(cookie, TF_DEFAULT_SELECTION, &mut selection, &mut fetched) }.ok()?;
    if fetched == 0 {
        return None;
    }
    unsafe { ManuallyDrop::take(&mut selection[0].range) }
}

fn select(context: &ITfContext, cookie: u32, range: &ITfRange) -> Result<()> {
    let selection = TF_SELECTION {
        range: ManuallyDrop::new(Some(range.clone())),
        style: TF_SELECTIONSTYLE { ase: TF_AE_NONE, fInterimChar: FALSE },
    };
    let result = unsafe { context.SetSelection(cookie, &[selection.clone()]) };
    drop(ManuallyDrop::into_inner(selection.range));
    result
}

thread_local! {
    /// The widths of a filler and of a joint (space.rs) by window, as the field drew them.
    static SPACE_WIDTHS: RefCell<HashMap<isize, (f64, f64)>> = RefCell::new(HashMap::new());
    /// The height of a line of text by window, learnt from the caret where the line had text
    /// before it (physical pixels). An empty line reports its whole line box instead (Chrome,
    /// Electron), which would make the formula too large: VisualTeX takes this height there (as
    /// the macOS input method does). Kept across activations (the service is made anew each
    /// time); by window, so that one field's text (a search box's) does not size another's.
    static TEXT_HEIGHTS: RefCell<HashMap<isize, i32>> = RefCell::new(HashMap::new());
}

/// Text stands before `range` on its line.
fn text_before(cookie: u32, range: &ITfRange) -> bool {
    let Ok(before) = (unsafe { range.Clone() }) else { return false };
    let mut moved = 0i32;
    let mut text = [0u16; 1];
    let mut read = 0u32;
    unsafe {
        if before.Collapse(cookie, TF_ANCHOR_START).is_err()
            || before.ShiftStart(cookie, -1, &mut moved, std::ptr::null()).is_err()
            || moved != -1
            || before.GetText(cookie, 0, &mut text, &mut read).is_err()
        {
            return false;
        }
    }
    read == 1 && !matches!(text[0], 0x0a | 0x0b | 0x0c | 0x0d | 0x2028 | 0x2029)
}

/// The text-height fields of a geometry message. `learn`: the rectangle is the caret's, so where
/// text stands before it its height is the text's (never the formula's anchor, whose height some
/// fields make up, nor a clipped rectangle).
fn add_text_height(message: &mut Value, cookie: u32, range: &ITfRange, extent: &Extent, learn: bool) {
    let before = text_before(cookie, range);
    let height = extent.rect.bottom - extent.rect.top;
    if learn && before && !extent.clipped && (4..=400).contains(&height) {
        TEXT_HEIGHTS.with(|heights| heights.borrow_mut().insert(extent.window, height));
    }
    message["textBefore"] = json!(before);
    message["clipped"] = json!(extent.clipped);
    if let Some(learnt) = TEXT_HEIGHTS.with(|heights| heights.borrow().get(&extent.window).copied()) {
        message["textHeight"] = json!(learnt);
    }
}

/// One line of the log for a rectangle reported to VisualTeX (`source`: where an IMM32 caret came
/// from, see field_extent).
fn log_extent(what: &str, extent: &Extent, message: &Value, source: Option<&str>) {
    let rect = &extent.rect;
    log::write(&format!(
        "{what} {},{} {}x{} clipped={} window={:x} textBefore={} textHeight={}{}",
        rect.left,
        rect.top,
        rect.right - rect.left,
        rect.bottom - rect.top,
        extent.clipped,
        extent.window,
        message.get("textBefore").and_then(Value::as_bool).unwrap_or(false),
        message.get("textHeight").and_then(Value::as_i64).unwrap_or(0),
        source.map(|source| format!(" source={source}")).unwrap_or_default(),
    ));
}

fn rect_payload(rect: &RECT) -> Value {
    json!([rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top])
}

/// A caret-like rectangle at a rectangle's left edge.
fn caret_payload(edges: Edges, offset: i32) -> Value {
    json!([edges.0 + offset, edges.1, 1, edges.3 - edges.1])
}

/// `length` characters of `range` from its `start`-th.
fn sub_range(range: &ITfRange, cookie: u32, start: usize, length: usize) -> Option<ITfRange> {
    let part = unsafe { range.Clone() }.ok()?;
    let mut moved = 0i32;
    unsafe {
        part.Collapse(cookie, TF_ANCHOR_START).ok()?;
        part.ShiftEnd(cookie, (start + length) as i32, &mut moved, std::ptr::null()).ok()?;
        if moved != (start + length) as i32 {
            return None;
        }
        part.ShiftStart(cookie, start as i32, &mut moved, std::ptr::null()).ok()?;
    }
    Some(part)
}

/// The screen rectangle of the `index`-th character of `range` (physical pixels).
fn char_rect(context: &ITfContext, cookie: u32, range: &ITfRange, index: usize) -> Option<RECT> {
    let part = sub_range(range, cookie, index, 1)?;
    range_rect(context, cookie, &part).filter(|rect| rect.right >= rect.left && rect.bottom - rect.top < 2000)
}

/// Two rectangles on one line of the field.
fn same_line(a: &RECT, b: &RECT) -> bool {
    (a.bottom - b.bottom).abs() * 2 < (a.bottom - a.top).max(1)
}

/// A filler's and a joint's width before the field has shown them: from the height of a line of
/// its text (a blank is about half an em, a word space a quarter).
fn estimated_space_widths(window: isize, line_height: i32) -> (f64, f64) {
    if let Some(known) = SPACE_WIDTHS.with(|widths| widths.borrow().get(&window).copied()) {
        return known;
    }
    let height = TEXT_HEIGHTS.with(|heights| heights.borrow().get(&window).copied()).unwrap_or(line_height).clamp(8, 200) as f64;
    let em = height / 1.32;
    (em * 0.5, em * 0.27)
}

// MARK: Pinyin in the field

fn set_attribute(shared: &Shared, context: &ITfContext, cookie: u32, range: &ITfRange) {
    set_attribute_atom(shared.display_atom.get(), context, cookie, range);
}

/// Gives `range` the display attribute registered as `atom` (0: none).
fn set_attribute_atom(atom: u32, context: &ITfContext, cookie: u32, range: &ITfRange) {
    if atom == 0 {
        return;
    }
    if let Ok(property) = unsafe { context.GetProperty(&GUID_PROP_ATTRIBUTE) } {
        let value = VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                    vt: VT_I4,
                    wReserved1: 0,
                    wReserved2: 0,
                    wReserved3: 0,
                    Anonymous: VARIANT_0_0_0 { lVal: atom as i32 },
                }),
            },
        };
        unsafe {
            let _ = property.SetValue(cookie, range, &value);
        }
    }
}

fn clear_attribute(context: &ITfContext, cookie: u32, range: &ITfRange) {
    if let Ok(property) = unsafe { context.GetProperty(&GUID_PROP_ATTRIBUTE) } {
        unsafe {
            let _ = property.Clear(cookie, range);
        }
    }
}

/// RIME's commit into the field, its spelling as the field's composition, its candidates under it.
fn apply_pinyin(shared: &Rc<Shared>, context: &ITfContext, reply: Reply) {
    let state = shared.clone();
    let writer = context.clone();
    let result = edit::request_now(context, shared.client.get(), move |cookie| {
        if let Some(text) = &reply.commit {
            let existing = state.pinyin.borrow_mut().take();
            state.pinyin_context.borrow_mut().take();
            match existing {
                Some(composition) => {
                    let range = unsafe { composition.GetRange()? };
                    clear_attribute(&writer, cookie, &range);
                    unsafe {
                        range.SetText(cookie, 0, &utf16(text))?;
                        range.Collapse(cookie, TF_ANCHOR_END)?;
                    }
                    select(&writer, cookie, &range)?;
                    unsafe { composition.EndComposition(cookie)? };
                }
                None => {
                    let insert: ITfInsertAtSelection = writer.cast()?;
                    let range = unsafe { insert.InsertTextAtSelection(cookie, INSERT_TEXT_AT_SELECTION_FLAGS(0), &utf16(text))? };
                    unsafe { range.Collapse(cookie, TF_ANCHOR_END)? };
                    select(&writer, cookie, &range)?;
                }
            }
        }
        if !reply.preedit.is_empty() {
            let existing = state.pinyin.borrow().clone();
            let mut before: Option<RECT> = None;
            let composition = match existing {
                Some(composition) => composition,
                None => {
                    let insert: ITfInsertAtSelection = writer.cast()?;
                    let range = unsafe { insert.InsertTextAtSelection(cookie, TF_IAS_QUERYONLY, &[])? };
                    before = range_rect(&writer, cookie, &range);
                    let composer: ITfContextComposition = writer.cast()?;
                    let sink = state.composition_sink.borrow().clone();
                    let composition = unsafe { composer.StartComposition(cookie, &range, sink.as_ref())? };
                    *state.pinyin.borrow_mut() = Some(composition.clone());
                    *state.pinyin_context.borrow_mut() = Some(writer.clone());
                    composition
                }
            };
            let range = unsafe { composition.GetRange()? };
            unsafe { range.SetText(cookie, 0, &utf16(&reply.preedit))? };
            set_attribute(&state, &writer, cookie, &range);
            let caret = unsafe { range.Clone()? };
            let mut moved = 0i32;
            unsafe {
                caret.Collapse(cookie, TF_ANCHOR_START)?;
                caret.ShiftEnd(cookie, reply.cursor as i32, &mut moved, std::ptr::null())?;
                caret.Collapse(cookie, TF_ANCHOR_END)?;
            }
            select(&writer, cookie, &caret)?;
            // The candidates under the spelling's start (where the caret was until the field
            // lays the composition out).
            let at = range_rect(&writer, cookie, &range).or(before).map(|rect| edges(&rect)).or(state.placed.get());
            if let Some(at) = at {
                state.placed.set(Some(at));
                state.engine.call(json!({ "op": "show", "caret": caret_payload(at, 0) }));
            }
        } else if let Some(composition) = state.pinyin.borrow_mut().take() {
            state.pinyin_context.borrow_mut().take();
            let range = unsafe { composition.GetRange()? };
            unsafe {
                range.SetText(cookie, 0, &[])?;
                composition.EndComposition(cookie)?;
            }
        }
        if reply.preedit.is_empty() {
            state.placed.set(None);
            state.engine.call_existing(json!({ "op": "hide" }));
        }
        Ok(())
    });
    if let Err(error) = result {
        log::write(&format!("pinyin edit failed: {error}"));
    }
    if shared.pinyin.borrow().is_some() {
        shared.place_polls.set(0);
        unsafe {
            SetTimer(Some(shared.window()), PLACE_TIMER, 30, None);
        }
    }
}

/// Follows the pinyin composition with the candidates (the field lays it out late).
fn place_candidates(shared: &Rc<Shared>) {
    let polls = shared.place_polls.get() + 1;
    shared.place_polls.set(polls);
    let (Some(context), Some(composition)) = (shared.pinyin_context.borrow().clone(), shared.pinyin.borrow().clone()) else {
        unsafe {
            let _ = KillTimer(Some(shared.window()), PLACE_TIMER);
        }
        return;
    };
    if polls >= PLACE_POLLS {
        unsafe {
            let _ = KillTimer(Some(shared.window()), PLACE_TIMER);
        }
    }
    let state = shared.clone();
    let reader = context.clone();
    let _ = edit::request(&context, shared.client.get(), move |cookie| {
        let Ok(range) = (unsafe { composition.GetRange() }) else { return Ok(()) };
        let Some(rect) = range_rect(&reader, cookie, &range) else { return Ok(()) };
        let at = edges(&rect);
        if state.placed.get() != Some(at) {
            state.placed.set(Some(at));
            state.engine.call(json!({ "op": "show", "caret": caret_payload(at, 0) }));
        }
        Ok(())
    });
}

/// What is being spelled goes into the field as it is spelled (a formula opens).
fn finish_pinyin(shared: &Rc<Shared>) {
    let (Some(context), Some(_)) = (shared.pinyin_context.borrow().clone(), shared.pinyin.borrow().clone()) else { return };
    let text = shared
        .engine
        .call(json!({ "op": "commit" }))
        .and_then(|reply| reply.get("commit").and_then(Value::as_str).map(str::to_string))
        .unwrap_or_default();
    apply_pinyin(shared, &context, Reply { handled: true, commit: Some(text).filter(|text| !text.is_empty()), ..Default::default() });
    // Nothing committed: the composition ends with what it shows.
    if let Some(composition) = shared.pinyin.borrow_mut().take() {
        shared.pinyin_context.borrow_mut().take();
        let _ = edit::request(&context, shared.client.get(), move |cookie| unsafe { composition.EndComposition(cookie) });
    }
}

/// The field lost the composition (focus moved, the field ended it): RIME forgets the spelling.
fn drop_pinyin(shared: &Rc<Shared>, end_composition: bool) {
    let composition = shared.pinyin.borrow_mut().take();
    let context = shared.pinyin_context.borrow_mut().take();
    shared.placed.set(None);
    shared.engine.call_existing(json!({ "op": "clear" }));
    if let (true, Some(composition), Some(context)) = (end_composition, composition, context) {
        let _ = edit::request(&context, shared.client.get(), move |cookie| unsafe { composition.EndComposition(cookie) });
    }
}

/// "中" / "英" for a moment when Shift switched (under `at`, else under the caret).
fn note_ascii(shared: &Rc<Shared>, ascii: bool, at: Option<Edges>) {
    if shared.ascii.replace(ascii) == ascii {
        return;
    }
    let text = if ascii { "英" } else { "中" };
    if let Some(at) = at {
        shared.engine.notify(json!({ "op": "indicator", "text": text, "caret": caret_payload(at, 0) }));
        return;
    }
    let Some(context) = shared.focused_context() else { return };
    let state = shared.clone();
    let reader = context.clone();
    let _ = edit::request(&context, shared.client.get(), move |cookie| {
        if let Some(rect) = selection_range(&reader, cookie).and_then(|range| range_rect(&reader, cookie, &range)) {
            state.engine.notify(json!({ "op": "indicator", "text": text, "caret": caret_payload(edges(&rect), 0) }));
        }
        Ok(())
    });
}

// MARK: The formula

/// Reports the caret (or, with a formula open, the anchor) to VisualTeX.
fn report(shared: &Rc<Shared>, kind: &'static str) {
    let Some(context) = shared.focused_context() else {
        shared.send(json!({ "t": kind, "client": client_name() }));
        return;
    };
    let state = shared.clone();
    let reader = context.clone();
    let _ = edit::request(&context, shared.client.get(), move |cookie| {
        let composition_range = state.formula.borrow().as_ref().and_then(|composition| unsafe { composition.GetRange() }.ok());
        // The caret's own rectangle (no formula open) teaches the text's height.
        let learn = composition_range.is_none();
        let range = composition_range.or_else(|| selection_range(&reader, cookie));
        let (extent, source) = field_extent(&reader, range.as_ref().and_then(|range| range_extent(&reader, cookie, range)));
        let mut message = json!({ "t": kind, "client": client_name() });
        if let Some(extent) = extent {
            message["caret"] = rect_payload(&extent.rect);
            message["line"] = rect_payload(&extent.rect);
            match range.as_ref() {
                Some(range) => add_text_height(&mut message, cookie, range, &extent, learn),
                None => message["clipped"] = json!(false),
            }
            if source.is_some() {
                message["source"] = json!("imm");
            }
            log_extent(kind, &extent, &message, source.as_deref());
        }
        state.send(message);
        Ok(())
    });
}

/// Starts the formula's composition at the caret: the anchor, selected.
fn start_formula(shared: &Rc<Shared>) {
    let Some(context) = shared.focused_context() else {
        log::write("session on without a focused field");
        return;
    };
    let state = shared.clone();
    let writer = context.clone();
    let result = edit::request(&context, shared.client.get(), move |cookie| {
        let transitory = types_through_imm32(&writer);
        let insert: ITfInsertAtSelection = writer.cast()?;
        let range = unsafe { insert.InsertTextAtSelection(cookie, TF_IAS_QUERYONLY, &[])? };
        let composer: ITfContextComposition = writer.cast()?;
        let sink = state.composition_sink.borrow().clone();
        let composition = unsafe { composer.StartComposition(cookie, &range, sink.as_ref())? };
        let range = unsafe { composition.GetRange()? };
        if transitory {
            // a field typing through IMM32 reports made-up rectangles for parts of the
            // composition: the formula stays on one line there, over a zero-width anchor
            unsafe { range.SetText(cookie, 0, IMM_ANCHOR)? };
            set_attribute_atom(state.anchor_atom.get(), &writer, cookie, &range);
            select(&writer, cookie, &range)?;
        } else {
            let caret = range_rect(&writer, cookie, &range);
            let window = view_window(&writer);
            let (filler, joint) = estimated_space_widths(window, caret.map(|rect| rect.bottom - rect.top).unwrap_or(20));
            state.filler_width.set(filler);
            state.joint_width.set(joint);
            let held = space::layout(&[], filler, joint);
            unsafe { range.SetText(cookie, 0, &held.text)? };
            set_attribute_atom(state.space_atom.get(), &writer, cookie, &range);
            if let Some(anchor) = sub_range(&range, cookie, 0, 1) {
                select(&writer, cookie, &anchor)?;
            }
            *state.space.borrow_mut() = held;
        }
        *state.formula.borrow_mut() = Some(composition);
        *state.formula_context.borrow_mut() = Some(writer.clone());
        state.formula_transitory.set(transitory);
        state.last_rect.set(None);
        state.stable_rect.set(None);
        state.last_pieces.borrow_mut().clear();
        *state.reported_pieces.borrow_mut() = None;
        state.calibrations.set(0);
        state.polls.set(0);
        unsafe {
            SetTimer(Some(state.window()), POLL_TIMER, 16, None);
        }
        log::write(if transitory { "formula composition started (IMM32 field)" } else { "formula composition started" });
        Ok(())
    });
    if let Err(error) = result {
        log::write(&format!("formula composition failed: {error}"));
    }
}

/// The field holds the formula's pieces (space.rs) and their places are followed anew.
fn hold(shared: &Rc<Shared>) {
    let (Some(context), Some(composition)) = (shared.formula_context.borrow().clone(), shared.formula.borrow().clone()) else { return };
    if shared.formula_transitory.get() {
        return;
    }
    let held = space::layout(&shared.piece_widths.borrow(), shared.filler_width.get(), shared.joint_width.get());
    if *shared.space.borrow() == held {
        return;
    }
    let state = shared.clone();
    let writer = context.clone();
    let result = edit::request(&context, shared.client.get(), move |cookie| {
        let range = unsafe { composition.GetRange()? };
        unsafe { range.SetText(cookie, 0, &held.text)? };
        set_attribute_atom(state.space_atom.get(), &writer, cookie, &range);
        if let Some(anchor) = sub_range(&range, cookie, 0, 1) {
            select(&writer, cookie, &anchor)?;
        }
        *state.space.borrow_mut() = held.clone();
        state.last_pieces.borrow_mut().clear();
        Ok(())
    });
    if let Err(error) = result {
        log::write(&format!("formula space not held: {error}"));
        return;
    }
    shared.polls.set(0);
    unsafe {
        SetTimer(Some(shared.window()), POLL_TIMER, 16, None);
    }
}

/// Where the field put each piece (the first character of each): reported once two reads agree,
/// with where the held space starts and ends; the filler's and the joint's widths are read off the
/// same characters (true: the space was held again with them).
fn follow_pieces(state: &Rc<Shared>, reader: &ITfContext, cookie: u32, range: &ITfRange, polls: u32) -> bool {
    let held = state.space.borrow().clone();
    if held.starts.is_empty() {
        return false;
    }
    let mut rects: Vec<RECT> = Vec::with_capacity(held.starts.len());
    for &start in &held.starts {
        let Some(rect) = char_rect(reader, cookie, range, start) else { break };
        // each piece after the one before on its line, or on a lower line: a field that answers
        // one rectangle for every part of the composition places nothing
        if let Some(before) = rects.last() {
            let onward = if same_line(before, &rect) { rect.left > before.left } else { rect.bottom > before.bottom };
            if !onward {
                break;
            }
        }
        rects.push(rect);
    }
    let current: Vec<Edges> = rects.iter().map(edges).collect();
    let previous = std::mem::replace(&mut *state.last_pieces.borrow_mut(), current.clone());
    if current.is_empty() || previous != current {
        return false;
    }
    let first = rects[0];
    let place = |rect: &RECT, x: i32| json!([x - first.left, rect.bottom - first.bottom]);
    let places: Vec<Value> = rects.iter().map(|rect| place(rect, rect.left)).collect();
    let mut message = json!({ "t": "pieces", "places": places, "count": held.starts.len() });
    if let Some(anchor) = char_rect(reader, cookie, range, 0).filter(|anchor| same_line(anchor, &first)) {
        message["start"] = json!([anchor.left - first.left, 0]);
    }
    if rects.len() == held.starts.len() {
        // after the last character, or (a field that answers nothing for it) the last piece's start
        // and its fillers
        match held.text.len().checked_sub(1).and_then(|index| char_rect(reader, cookie, range, index)) {
            Some(last) => message["end"] = place(&last, last.right),
            None => {
                if let (Some(start), Some(&count)) = (rects.last(), held.counts.last()) {
                    // estimated (a third component): VisualTeX covers no further than its formula's end
                    let x = start.left + (count as f64 * state.filler_width.get()).round() as i32;
                    message["end"] = json!([x - first.left, start.bottom - first.bottom, 1]);
                }
            }
        }
    }
    if state.reported_pieces.borrow().as_ref() != Some(&message) {
        let lines = rects.windows(2).filter(|pair| !same_line(&pair[0], &pair[1])).count() + 1;
        log::write(&format!("pieces {}/{} lines={lines} polls={polls}", rects.len(), held.starts.len()));
        state.send(message.clone());
        *state.reported_pieces.borrow_mut() = Some(message);
    }
    // the widths as the field draws them: a filler over the longest piece on one line (a field may
    // round each character's place to a pixel), a joint between a piece and the next on its line
    let mut filler: Option<(usize, f64)> = None;
    let mut joint: Option<f64> = None;
    for (index, &start) in held.starts.iter().enumerate().take(rects.len()) {
        let count = held.counts[index];
        let Some(last) = char_rect(reader, cookie, range, start + count - 1) else { continue };
        if !same_line(&last, &rects[index]) || last.right <= rects[index].left {
            continue;
        }
        // a run long enough that a field's pixel rounding is a small part of it
        if count >= 6 && filler.map(|(known, _)| count > known).unwrap_or(true) {
            filler = Some((count, (last.right - rects[index].left) as f64 / count as f64));
        }
        if joint.is_none() && index + 1 < rects.len() && same_line(&rects[index + 1], &last) && rects[index + 1].left >= last.right {
            joint = Some((rects[index + 1].left - last.right) as f64);
        }
    }
    let window = view_window(reader);
    let mut changed = false;
    if let Some((count, width)) = filler.filter(|(_, width)| (0.5..60.0).contains(width)) {
        if (width - state.filler_width.get()).abs() * count as f64 > 1.5 {
            log::write(&format!("filler {:.2} -> {width:.2} from {count}", state.filler_width.get()));
            state.filler_width.set(width);
            changed = true;
        }
    }
    if let Some(width) = joint.filter(|width| (0.0..60.0).contains(width)) {
        if (width - state.joint_width.get()).abs() > 1.5 {
            log::write(&format!("joint {:.2} -> {width:.2}", state.joint_width.get()));
            state.joint_width.set(width);
            changed = true;
        }
    }
    if filler.is_some() || joint.is_some() {
        SPACE_WIDTHS.with(|widths| widths.borrow_mut().insert(window, (state.filler_width.get(), state.joint_width.get())));
    }
    if changed && state.calibrations.get() < 3 {
        state.calibrations.set(state.calibrations.get() + 1);
        return true;
    }
    false
}

/// Polls the anchor's rectangle: reported once two reads agree, then followed while the formula
/// is open (scrolling, a late layout).
fn poll_anchor(shared: &Rc<Shared>) {
    let (Some(context), Some(composition)) = (shared.formula_context.borrow().clone(), shared.formula.borrow().clone()) else {
        stop_polling(shared);
        return;
    };
    if shared.rehold.replace(false) {
        *shared.space.borrow_mut() = space::Layout::default();
        hold(shared);
    }
    let polls = shared.polls.get() + 1;
    shared.polls.set(polls);
    if polls == FAST_POLLS {
        unsafe {
            SetTimer(Some(shared.window()), POLL_TIMER, 250, None);
        }
    }
    let state = shared.clone();
    let reader = context.clone();
    let _ = edit::request(&context, shared.client.get(), move |cookie| {
        let Ok(range) = (unsafe { composition.GetRange() }) else { return Ok(()) };
        // the formula's start: its first piece's (the field may have put it on the next line)
        let start = if state.formula_transitory.get() {
            None
        } else {
            state.space.borrow().starts.first().and_then(|&first| sub_range(&range, cookie, first, 1))
        };
        let (extent, source) = field_extent(&reader, range_extent(&reader, cookie, start.as_ref().unwrap_or(&range)));
        let Some(extent) = extent else { return Ok(()) };
        let rect = extent.rect;
        let current = edges(&rect);
        let previous = state.last_rect.replace(Some(current));
        if previous == Some(current) && state.stable_rect.get() != Some(current) {
            state.stable_rect.set(Some(current));
            let mut anchor = rect;
            anchor.right = anchor.left;
            let mut message = json!({
                "t": "geometry",
                "client": client_name(),
                "caret": rect_payload(&anchor),
                "line": rect_payload(&rect),
            });
            add_text_height(&mut message, cookie, &range, &extent, false);
            if source.is_some() {
                message["source"] = json!("imm");
            }
            log_extent(&format!("anchor polls={polls}"), &extent, &message, source.as_deref());
            state.send(message);
        }
        if !state.formula_transitory.get() && follow_pieces(&state, &reader, cookie, &range, polls) {
            // the widths changed: the space is held again at the next poll (outside this session)
            state.rehold.set(true);
        }
        Ok(())
    });
}

fn stop_polling(shared: &Shared) {
    unsafe {
        let _ = KillTimer(Some(shared.window()), POLL_TIMER);
    }
}

/// Ends the formula's composition with `text` in its place (empty: removed).
fn finish_formula(shared: &Rc<Shared>, text: String, reply: Option<&'static str>) {
    stop_polling(shared);
    shared.composing.set(false);
    let composition = shared.formula.borrow_mut().take();
    let context = shared.formula_context.borrow_mut().take();
    let (Some(composition), Some(context)) = (composition, context) else {
        // The composition never started or the field dropped it unseen: a formula committed then
        // still goes in at the field's caret (it used to vanish, the field taking no input).
        if !text.is_empty() {
            if let Some(context) = shared.focused_context() {
                let writer = context.clone();
                let inserted = edit::request(&context, shared.client.get(), move |cookie| {
                    let insert: ITfInsertAtSelection = writer.cast()?;
                    let range = unsafe { insert.InsertTextAtSelection(cookie, INSERT_TEXT_AT_SELECTION_FLAGS(0), &utf16(&text))? };
                    unsafe { range.Collapse(cookie, TF_ANCHOR_END)? };
                    let _ = select(&writer, cookie, &range);
                    Ok(())
                });
                log::write(&format!("formula inserted without its composition: {}", inserted.is_ok()));
            }
        }
        if let Some(reply) = reply {
            shared.send(json!({ "t": reply }));
        }
        return;
    };
    let state = shared.clone();
    let writer = context.clone();
    let result = edit::request(&context, shared.client.get(), move |cookie| {
        let range = unsafe { composition.GetRange()? };
        // the formula's text is no clause being converted
        clear_attribute(&writer, cookie, &range);
        unsafe {
            range.SetText(cookie, 0, &utf16(&text))?;
            range.Collapse(cookie, TF_ANCHOR_END)?;
        }
        let _ = select(&writer, cookie, &range);
        unsafe { composition.EndComposition(cookie)? };
        if let Some(reply) = reply {
            state.send(json!({ "t": reply }));
        }
        Ok(())
    });
    if let Err(error) = result {
        log::write(&format!("formula commit failed: {error}"));
    }
}

/// A formula opens in English (its letters are variables); Shift switches to Chinese for text.
/// The input mode comes back after.
fn set_formula_input_mode(shared: &Rc<Shared>, on: bool) {
    if on {
        if shared.ascii_before_formula.get().is_none() {
            let ascii = shared
                .engine
                .call(json!({ "op": "ascii" }))
                .and_then(|reply| reply.get("ascii").and_then(Value::as_bool))
                .unwrap_or(false);
            shared.ascii_before_formula.set(Some(ascii));
        }
        shared.engine.call(json!({ "op": "ascii", "on": true }));
        shared.ascii.set(true);
    } else if let Some(ascii) = shared.ascii_before_formula.take() {
        shared.engine.call(json!({ "op": "ascii", "on": ascii }));
        shared.ascii.set(ascii);
    }
}

/// Pinyin in the formula's text: RIME's spelling is the editor's composition (shown in the
/// formula), its commit ends the composition with the chosen characters, its candidates go
/// under the formula's caret.
fn formula_pinyin(shared: &Rc<Shared>, reply: &Reply) {
    // the formula's caret: from the formula's start, on its line (a later one when it wraps)
    let (dx, dy) = (shared.formula_caret_x.get(), shared.formula_caret_y.get());
    let caret = shared.stable_rect.get().map(|edges| (edges.0 + dx, edges.1 + dy, edges.0 + dx + 1, edges.3 + dy));
    note_ascii(shared, reply.ascii, caret);
    if let Some(commit) = &reply.commit {
        shared.send(json!({ "t": "compose", "phase": "end", "text": commit }));
        shared.formula_pinyin.set(false);
    }
    if !reply.preedit.is_empty() {
        shared.send(json!({ "t": "compose", "phase": "update", "text": reply.preedit }));
        shared.formula_pinyin.set(true);
        if let Some(caret) = caret {
            shared.engine.call(json!({ "op": "show", "caret": caret_payload(caret, 0) }));
        }
    } else {
        if shared.formula_pinyin.replace(false) && reply.commit.is_none() {
            shared.send(json!({ "t": "compose", "phase": "end", "text": "" }));
        }
        shared.engine.call_existing(json!({ "op": "hide" }));
    }
}

/// Drops pinyin being spelled in the formula (the formula closed).
fn drop_formula_pinyin(shared: &Rc<Shared>) {
    shared.engine.call_existing(json!({ "op": "clear" }));
    if shared.formula_pinyin.replace(false) {
        shared.send(json!({ "t": "compose", "phase": "end", "text": "" }));
    }
}

fn close_formula(shared: &Rc<Shared>, text: String, reply: Option<&'static str>, return_to: Option<&Value>) {
    drop_formula_pinyin(shared);
    set_formula_input_mode(shared, false);
    finish_formula(shared, text, reply);
    if let Some(profile) = return_to.filter(|profile| profile.is_object()) {
        *shared.hand_back.borrow_mut() = Some(profile.clone());
        unsafe {
            let _ = PostMessageW(Some(shared.window()), WM_HAND_BACK, WPARAM(0), LPARAM(0));
        }
    }
}

/// A message from VisualTeX, on the UI thread.
fn receive(shared: &Rc<Shared>, message: &Value) {
    let Some(kind) = message.get("t").and_then(Value::as_str) else { return };
    match kind {
        "session" => {
            let on = message.get("on").and_then(Value::as_bool).unwrap_or(false);
            if on && !shared.composing.get() {
                finish_pinyin(shared);
                set_formula_input_mode(shared, true);
                shared.composing.set(true);
                shared.formula_text.borrow_mut().clear();
                shared.formula_caret_x.set(0);
                shared.formula_caret_y.set(0);
                shared.piece_widths.borrow_mut().clear();
                shared.calibrations.set(0);
                *shared.space.borrow_mut() = space::Layout::default();
                start_formula(shared);
            } else if !on && shared.composing.get() {
                close_formula(shared, String::new(), None, None);
            }
            log::write(&format!("session {}", if on { "on" } else { "off" }));
        }
        "inline" => {
            *shared.formula_text.borrow_mut() = message.get("text").and_then(Value::as_str).unwrap_or_default().to_string();
            shared.formula_caret_x.set(message.get("caretX").and_then(Value::as_f64).unwrap_or(0.0).round() as i32);
            shared.formula_caret_y.set(message.get("caretY").and_then(Value::as_f64).unwrap_or(0.0).round() as i32);
            if let Some(pieces) = message.get("pieces").and_then(Value::as_array) {
                let widths: Vec<f64> = pieces.iter().map(|piece| piece.as_f64().unwrap_or(0.0).clamp(0.0, 16384.0)).collect();
                if shared.composing.get() && *shared.piece_widths.borrow() != widths {
                    *shared.piece_widths.borrow_mut() = widths;
                    hold(shared);
                }
            }
        }
        "commit" => {
            let text = message.get("text").and_then(Value::as_str).unwrap_or_default().to_string();
            log::write(&format!("commit {} chars", text.chars().count()));
            close_formula(shared, text, Some("committed"), message.get("returnTo"));
        }
        "cancel" => {
            log::write("cancel");
            close_formula(shared, String::new(), Some("cancelled"), message.get("returnTo"));
        }
        "caret" => report(shared, "geometry"),
        "hold" => shared.hold.set(message.get("on").and_then(Value::as_bool).unwrap_or(false)),
        _ => {}
    }
}

// MARK: The window that receives VisualTeX's messages and the timers.

fn window_class_registered() {
    use std::sync::Once;
    static REGISTER: Once = Once::new();
    REGISTER.call_once(|| {
        let class = WNDCLASSW {
            lpfnWndProc: Some(window_proc),
            hInstance: crate::dll_instance(),
            lpszClassName: w!("VisualTeXTipLink"),
            ..Default::default()
        };
        unsafe { RegisterClassW(&class) };
    });
}

fn shared_of(window: HWND) -> Option<Rc<Shared>> {
    let pointer = unsafe { GetWindowLongPtrW(window, GWLP_USERDATA) } as *const Shared;
    if pointer.is_null() {
        return None;
    }
    Some(unsafe {
        Rc::increment_strong_count(pointer);
        Rc::from_raw(pointer)
    })
}

extern "system" fn window_proc(window: HWND, message: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match message {
        WM_LINK_MESSAGE => {
            let text = unsafe { Box::from_raw(lparam.0 as *mut String) };
            if let Some(shared) = shared_of(window) {
                if let Ok(value) = serde_json::from_str::<Value>(&text) {
                    receive(&shared, &value);
                }
            }
            LRESULT(0)
        }
        WM_LINK_CLOSED => {
            if let Some(shared) = shared_of(window) {
                shared.link.borrow_mut().take();
                if shared.composing.get() {
                    close_formula(&shared, String::new(), None, None);
                }
                log::write("link closed");
            }
            LRESULT(0)
        }
        WM_HAND_BACK => {
            if let Some(shared) = shared_of(window) {
                if let Some(profile) = shared.hand_back.borrow_mut().take() {
                    profile::activate(&profile);
                }
            }
            LRESULT(0)
        }
        WM_TIMER => {
            if let Some(shared) = shared_of(window) {
                match wparam.0 {
                    POLL_TIMER => poll_anchor(&shared),
                    PLACE_TIMER => place_candidates(&shared),
                    LINK_TIMER => {
                        if shared.link.borrow().is_none() {
                            if let Some(link) = Link::connect(window) {
                                link.send(&json!({ "t": "hello", "version": env!("CARGO_PKG_VERSION") }));
                                *shared.link.borrow_mut() = Some(link);
                                report(&shared, "activate");
                            }
                        }
                    }
                    _ => {}
                }
            }
            LRESULT(0)
        }
        _ => unsafe { DefWindowProcW(window, message, wparam, lparam) },
    }
}

fn create_window(shared: &Rc<Shared>) -> Option<HWND> {
    window_class_registered();
    let window = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE(0),
            w!("VisualTeXTipLink"),
            w!(""),
            WINDOW_STYLE(0),
            0,
            0,
            0,
            0,
            Some(HWND_MESSAGE),
            None,
            Some(crate::dll_instance()),
            None,
        )
    }
    .ok()?;
    let pointer = Rc::into_raw(shared.clone());
    unsafe { SetWindowLongPtrW(window, GWLP_USERDATA, pointer as isize) };
    Some(window)
}

fn destroy_window(shared: &Shared) {
    let window = HWND(shared.window.replace(0) as *mut _);
    if window.0.is_null() {
        return;
    }
    let pointer = unsafe { SetWindowLongPtrW(window, GWLP_USERDATA, 0) } as *const Shared;
    unsafe {
        let _ = DestroyWindow(window);
    }
    if !pointer.is_null() {
        drop(unsafe { Rc::from_raw(pointer) });
    }
}

// MARK: COM objects

#[implement(ITfTextInputProcessorEx, ITfTextInputProcessor, ITfDisplayAttributeProvider)]
pub(crate) struct TextService {
    shared: RefCell<Rc<Shared>>,
}

impl TextService {
    pub(crate) fn new() -> Self {
        Self { shared: RefCell::new(Rc::new(Shared::default())) }
    }
}

impl ITfDisplayAttributeProvider_Impl for TextService_Impl {
    fn EnumDisplayAttributeInfo(&self) -> Result<IEnumTfDisplayAttributeInfo> {
        Ok(display::AttributeEnum::new().into())
    }

    fn GetDisplayAttributeInfo(&self, guid: *const GUID) -> Result<ITfDisplayAttributeInfo> {
        if guid.is_null() {
            return Err(E_INVALIDARG.into());
        }
        display::AttributeInfo::of(unsafe { &*guid }).map(Into::into).ok_or_else(|| E_INVALIDARG.into())
    }
}

impl ITfTextInputProcessor_Impl for TextService_Impl {
    fn Activate(&self, thread_mgr: Ref<'_, ITfThreadMgr>, client: u32) -> Result<()> {
        ITfTextInputProcessorEx_Impl::ActivateEx(self, thread_mgr, client, 0)
    }

    fn Deactivate(&self) -> Result<()> {
        let shared = self.shared.borrow().clone();
        if shared.composing.get() {
            close_formula(&shared, String::new(), None, None);
        }
        if shared.pinyin.borrow().is_some() {
            drop_pinyin(&shared, true);
        }
        shared.engine.call_existing(json!({ "op": "hide" }));
        shared.send(json!({ "t": "deactivate", "client": client_name() }));
        if let Some(manager) = shared.thread_mgr.borrow_mut().take() {
            if let Ok(keystrokes) = manager.cast::<ITfKeystrokeMgr>() {
                unsafe {
                    let _ = keystrokes.UnadviseKeyEventSink(shared.client.get());
                }
            }
            if let Ok(source) = manager.cast::<ITfSource>() {
                unsafe {
                    let _ = source.UnadviseSink(shared.thread_sink_cookie.get());
                }
            }
        }
        unsafe {
            let _ = KillTimer(Some(shared.window()), LINK_TIMER);
            let _ = KillTimer(Some(shared.window()), PLACE_TIMER);
        }
        shared.engine.close();
        shared.key_sink.borrow_mut().take();
        shared.composition_sink.borrow_mut().take();
        shared.link.borrow_mut().take();
        destroy_window(&shared);
        // A fresh state for a later activation of this object.
        *self.shared.borrow_mut() = Rc::new(Shared::default());
        log::write("deactivated");
        Ok(())
    }
}

impl ITfTextInputProcessorEx_Impl for TextService_Impl {
    fn ActivateEx(&self, thread_mgr: Ref<'_, ITfThreadMgr>, client: u32, _flags: u32) -> Result<()> {
        let manager = thread_mgr.ok()?.clone();
        let shared = self.shared.borrow().clone();
        *shared.thread_mgr.borrow_mut() = Some(manager.clone());
        shared.client.set(client);
        if let Some(window) = create_window(&shared) {
            shared.window.set(window.0 as isize);
        }
        let key_sink: ITfKeyEventSink = KeySink { shared: shared.clone() }.into();
        let keystrokes: ITfKeystrokeMgr = manager.cast()?;
        unsafe { keystrokes.AdviseKeyEventSink(client, &key_sink, true)? };
        *shared.key_sink.borrow_mut() = Some(key_sink);
        let focus_sink: ITfThreadMgrEventSink = FocusSink { shared: shared.clone() }.into();
        let source: ITfSource = manager.cast()?;
        let cookie = unsafe { source.AdviseSink(&ITfThreadMgrEventSink::IID, &focus_sink)? };
        shared.thread_sink_cookie.set(cookie);
        let composition_sink: ITfCompositionSink = CompositionSink { shared: shared.clone() }.into();
        *shared.composition_sink.borrow_mut() = Some(composition_sink);
        let categories: Result<ITfCategoryMgr> = unsafe { CoCreateInstance(&CLSID_TF_CategoryMgr, None, CLSCTX_INPROC_SERVER) };
        if let Ok(categories) = categories {
            if let Ok(atom) = unsafe { categories.RegisterGUID(&display::GUID_INPUT) } {
                shared.display_atom.set(atom);
            }
            if let Ok(atom) = unsafe { categories.RegisterGUID(&display::GUID_ANCHOR) } {
                shared.anchor_atom.set(atom);
            }
            if let Ok(atom) = unsafe { categories.RegisterGUID(&display::GUID_SPACE) } {
                shared.space_atom.set(atom);
            }
        }
        if let Some(link) = Link::connect(shared.window()) {
            link.send(&json!({ "t": "hello", "version": env!("CARGO_PKG_VERSION") }));
            *shared.link.borrow_mut() = Some(link);
        }
        unsafe {
            SetTimer(Some(shared.window()), LINK_TIMER, 1000, None);
        }
        log::write("activated");
        report(&shared, "activate");
        Ok(())
    }
}

#[implement(ITfKeyEventSink)]
struct KeySink {
    shared: Rc<Shared>,
}

impl KeySink {
    /// While a formula is open its keys are VisualTeX's (but the system's and the application's
    /// Control chords).
    fn takes_formula(&self, vk: u32) -> bool {
        self.shared.composing.get()
            && !keys::belongs_to_system(vk)
            && (rimekeys::is_shift(vk) || keys::belongs_to_formula(vk))
    }

    /// A key offered to RIME outside a formula: None when it is the application's.
    fn offer(&self, vk: u32, lparam: isize, release: bool) -> Option<Reply> {
        let mods = keys::modifiers();
        if mods.win {
            return None;
        }
        let spelling = self.shared.pinyin.borrow().is_some();
        if release {
            // Key-ups matter for Shift (it switches Chinese and English) and Control.
            if !(rimekeys::is_shift(vk) || rimekeys::is_control(vk)) {
                return None;
            }
        } else if !spelling {
            if rimekeys::is_modifier(vk) && !(rimekeys::is_shift(vk) || rimekeys::is_control(vk)) {
                return None;
            }
            // Control and Alt shortcuts are the application's while nothing is spelled.
            if (mods.ctrl || mods.alt) && !rimekeys::is_modifier(vk) {
                return None;
            }
        }
        let keysym = rimekeys::keysym(vk, lparam)?;
        self.shared.engine.key(keysym, rimekeys::mask(release))
    }

    fn take_pending(&self, vk: u32, release: bool) -> Option<Reply> {
        let pending = self.shared.pending.borrow_mut().take();
        match pending {
            Some((key, up, reply)) if key == vk && up == release => Some(reply),
            _ => None,
        }
    }

    fn test(&self, vk: u32, lparam: isize, release: bool) -> bool {
        if self.shared.composing.get() {
            return self.takes_formula(vk);
        }
        let reply = self.offer(vk, lparam, release);
        let handled = reply.as_ref().map(|reply| reply.handled).unwrap_or(false);
        *self.shared.pending.borrow_mut() = reply.map(|reply| (vk, release, reply));
        handled
    }

    fn key(&self, context: Option<&ITfContext>, vk: u32, lparam: isize, release: bool) -> bool {
        let shared = &self.shared;
        if shared.composing.get() {
            return self.formula_key(vk, lparam, release);
        }
        let Some(reply) = self.take_pending(vk, release).or_else(|| self.offer(vk, lparam, release)) else { return false };
        let spelling = shared.pinyin.borrow().is_some();
        if reply.handled || reply.commit.is_some() || (spelling && reply.preedit.is_empty()) {
            if let Some(context) = context {
                apply_pinyin(shared, context, reply.clone());
            }
        }
        note_ascii(shared, reply.ascii, shared.placed.get());
        reply.handled
    }

    /// A key while a formula is open: Shift to RIME (Chinese or English); Control and Alt chords
    /// to the formula (its shortcuts, undo, paste, rows); in Chinese the keys spell pinyin and
    /// the chosen characters become the formula's text; the rest is the formula's.
    fn formula_key(&self, vk: u32, lparam: isize, release: bool) -> bool {
        if !self.takes_formula(vk) {
            return false;
        }
        let shared = &self.shared;
        if rimekeys::is_shift(vk) {
            if let Some(reply) = rimekeys::keysym(vk, lparam).and_then(|keysym| shared.engine.key(keysym, rimekeys::mask(release))) {
                formula_pinyin(shared, &reply);
            }
            return true;
        }
        if release {
            return true;
        }
        let repeat = (lparam >> 30) & 1 == 1;
        let mods = keys::modifiers();
        if !(mods.ctrl || mods.alt) {
            if let Some(reply) = rimekeys::keysym(vk, lparam).and_then(|keysym| shared.engine.key(keysym, rimekeys::mask(false))) {
                if reply.handled {
                    formula_pinyin(shared, &reply);
                    return true;
                }
            }
        }
        shared.send(keys::message(vk, lparam, repeat));
        true
    }
}

impl ITfKeyEventSink_Impl for KeySink_Impl {
    fn OnSetFocus(&self, _foreground: BOOL) -> Result<()> {
        Ok(())
    }

    fn OnTestKeyDown(&self, _context: Ref<'_, ITfContext>, wparam: WPARAM, lparam: LPARAM) -> Result<BOOL> {
        Ok(BOOL::from(self.test(wparam.0 as u32, lparam.0, false)))
    }

    fn OnTestKeyUp(&self, _context: Ref<'_, ITfContext>, wparam: WPARAM, lparam: LPARAM) -> Result<BOOL> {
        Ok(BOOL::from(self.test(wparam.0 as u32, lparam.0, true)))
    }

    fn OnKeyDown(&self, context: Ref<'_, ITfContext>, wparam: WPARAM, lparam: LPARAM) -> Result<BOOL> {
        Ok(BOOL::from(self.key(context.as_ref(), wparam.0 as u32, lparam.0, false)))
    }

    fn OnKeyUp(&self, context: Ref<'_, ITfContext>, wparam: WPARAM, lparam: LPARAM) -> Result<BOOL> {
        Ok(BOOL::from(self.key(context.as_ref(), wparam.0 as u32, lparam.0, true)))
    }

    fn OnPreservedKey(&self, _context: Ref<'_, ITfContext>, _guid: *const GUID) -> Result<BOOL> {
        Ok(FALSE)
    }
}

#[implement(ITfThreadMgrEventSink)]
struct FocusSink {
    shared: Rc<Shared>,
}

impl ITfThreadMgrEventSink_Impl for FocusSink_Impl {
    fn OnInitDocumentMgr(&self, _document: Ref<'_, ITfDocumentMgr>) -> Result<()> {
        Ok(())
    }

    fn OnUninitDocumentMgr(&self, _document: Ref<'_, ITfDocumentMgr>) -> Result<()> {
        Ok(())
    }

    /// Focus moved to another field (or away): an open formula closes (the field it belonged to
    /// keeps nothing of it; VisualTeX keeps the draft for the next formula, as on macOS), pinyin
    /// being spelled stays as it is shown, VisualTeX hears of the new field.
    fn OnSetFocus(&self, focus: Ref<'_, ITfDocumentMgr>, _previous: Ref<'_, ITfDocumentMgr>) -> Result<()> {
        let shared = &self.shared;
        if shared.composing.get() && shared.hold.get() {
            // The screenshot selection: the field gets the keyboard back after it.
            return Ok(());
        }
        let left_formula = shared.composing.get();
        if left_formula {
            close_formula(shared, String::new(), None, None);
        }
        if shared.pinyin.borrow().is_some() {
            drop_pinyin(shared, true);
        }
        shared.engine.call_existing(json!({ "op": "hide" }));
        if left_formula || focus.is_none() {
            shared.send(json!({ "t": "deactivate", "client": client_name() }));
        }
        if focus.is_some() {
            report(shared, "activate");
        }
        Ok(())
    }

    fn OnPushContext(&self, _context: Ref<'_, ITfContext>) -> Result<()> {
        Ok(())
    }

    fn OnPopContext(&self, _context: Ref<'_, ITfContext>) -> Result<()> {
        Ok(())
    }
}

#[implement(ITfCompositionSink)]
struct CompositionSink {
    shared: Rc<Shared>,
}

impl ITfCompositionSink_Impl for CompositionSink_Impl {
    /// The field ended a composition itself (a click elsewhere in it): a formula's gets the
    /// formula's text, not the anchor; a pinyin spelling stays as shown.
    fn OnCompositionTerminated(&self, cookie: u32, composition: Ref<'_, ITfComposition>) -> Result<()> {
        let shared = &self.shared;
        let is_formula = match (composition.as_ref(), shared.formula.borrow().as_ref()) {
            (Some(ended), Some(formula)) => ended.as_raw() == formula.as_raw(),
            _ => false,
        };
        if is_formula && shared.composing.get() {
            stop_polling(shared);
            shared.composing.set(false);
            shared.formula.borrow_mut().take();
            shared.formula_context.borrow_mut().take();
            drop_formula_pinyin(shared);
            set_formula_input_mode(shared, false);
            if let Some(composition) = composition.as_ref() {
                if let Ok(range) = unsafe { composition.GetRange() } {
                    let text = utf16(&shared.formula_text.borrow());
                    if let Ok(context) = unsafe { range.GetContext() } {
                        clear_attribute(&context, cookie, &range);
                    }
                    unsafe {
                        let _ = range.SetText(cookie, 0, &text);
                    }
                }
            }
            shared.send(json!({ "t": "forced" }));
            log::write("formula composition terminated by the field");
        } else if shared.pinyin.borrow().is_some() {
            drop_pinyin(shared, false);
            shared.engine.call_existing(json!({ "op": "hide" }));
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn rect(left: i32, top: i32, width: i32, height: i32) -> RECT {
        RECT { left, top, right: left + width, bottom: top + height }
    }

    #[test]
    fn windows_made_up_rectangles_give_way_to_the_imm32_caret() {
        // WeChat at 150 %: Qt's cursor, 28 px; Windows' rectangles start at its bottom
        let caret = rect(2325, 1681, 2, 28);
        assert!(made_up(&rect(2325, 1709, 0, 20), &caret));
        assert!(made_up(&rect(2328, 1710, 0, 1), &caret));
        // the caret's own line, or nearly
        assert!(!made_up(&rect(2325, 1681, 0, 28), &caret));
        assert!(!made_up(&rect(2325, 1683, 0, 25), &caret));
        // a rectangle about another place stays
        assert!(!made_up(&rect(900, 400, 0, 20), &caret));
    }
}
