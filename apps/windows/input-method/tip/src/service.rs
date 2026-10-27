//! The text service: one per UI thread of an application that types with VisualTeX. VisualTeX is
//! a whole input method, as on the Mac: Chinese through pinyin (RIME, in VisualTeX Input) and
//! English (Shift switches), with formulas as one more mode. Its sinks share one state (`Shared`).
//!
//! Pinyin: the spelling is the field's composition (underlined), the candidates are VisualTeX
//! Input's window under it, and the chosen text replaces the composition.
//!
//! A formula (VisualTeX's "session on", from its shortcut): the field gets a composition holding
//! one selected hair space where the caret was (it takes almost no width and, being selected,
//! hides the field's caret; VisualTeX draws the formula there), every key but the system's goes
//! to VisualTeX (pinyin spelled in the formula's text goes to it as its composition), and
//! "commit" replaces the composition with the formula's text.

use crate::engine::{Engine, Reply};
use crate::link::{Link, WM_LINK_CLOSED, WM_LINK_MESSAGE};
use crate::{display, edit, keys, log, profile, rimekeys};
use serde_json::{json, Value};
use std::cell::{Cell, RefCell};
use std::mem::ManuallyDrop;
use std::rc::Rc;
use windows::core::{implement, w, Interface, Ref, Result, BOOL, GUID};
use windows::Win32::Foundation::{E_INVALIDARG, FALSE, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM};
use windows::Win32::System::Variant::{VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0, VT_I4};
use windows::Win32::UI::HiDpi::LogicalToPhysicalPointForPerMonitorDPI;
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};
use windows::Win32::UI::TextServices::{
    CLSID_TF_CategoryMgr, IEnumTfDisplayAttributeInfo, ITfCategoryMgr, ITfComposition, ITfCompositionSink,
    ITfCompositionSink_Impl, ITfContext, ITfContextComposition, ITfDisplayAttributeInfo, ITfDisplayAttributeProvider,
    ITfDisplayAttributeProvider_Impl, ITfDocumentMgr, ITfInsertAtSelection, ITfKeyEventSink, ITfKeyEventSink_Impl,
    ITfKeystrokeMgr, ITfRange, ITfSource, ITfTextInputProcessor, ITfTextInputProcessorEx,
    ITfTextInputProcessorEx_Impl, ITfTextInputProcessor_Impl, ITfThreadMgr, ITfThreadMgrEventSink,
    ITfThreadMgrEventSink_Impl, GUID_PROP_ATTRIBUTE, INSERT_TEXT_AT_SELECTION_FLAGS, TF_AE_NONE, TF_ANCHOR_END,
    TF_ANCHOR_START, TF_DEFAULT_SELECTION, TF_IAS_QUERYONLY, TF_SELECTION, TF_SELECTIONSTYLE,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetWindowLongPtrW, KillTimer,
    PostMessageW, RegisterClassW, SetTimer, SetWindowLongPtrW, GWLP_USERDATA, HWND_MESSAGE, WINDOW_EX_STYLE,
    WINDOW_STYLE, WM_APP, WM_TIMER, WNDCLASSW,
};

/// The formula's place in the field: one hair space, selected.
const ANCHOR: &[u16] = &[0x200A];
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
    /// What the formula inserts, for a field that ends the composition itself.
    formula_text: RefCell<String>,
    /// The formula's caret from its start, physical pixels.
    formula_caret_x: Cell<i32>,
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

/// The screen rectangle of `range` in physical pixels (top-left origin), whatever the
/// application's DPI awareness.
fn range_rect(context: &ITfContext, cookie: u32, range: &ITfRange) -> Option<RECT> {
    let view = unsafe { context.GetActiveView() }.ok()?;
    let mut rect = RECT::default();
    let mut clipped = BOOL(0);
    unsafe { view.GetTextExt(cookie, range, &mut rect, &mut clipped) }.ok()?;
    if rect.bottom <= rect.top {
        return None;
    }
    if let Ok(window) = unsafe { view.GetWnd() } {
        if !window.is_invalid() {
            let mut start = POINT { x: rect.left, y: rect.top };
            let mut end = POINT { x: rect.right, y: rect.bottom };
            unsafe {
                let _ = LogicalToPhysicalPointForPerMonitorDPI(Some(window), &mut start);
                let _ = LogicalToPhysicalPointForPerMonitorDPI(Some(window), &mut end);
            }
            rect = RECT { left: start.x, top: start.y, right: end.x, bottom: end.y };
        }
    }
    Some(rect)
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

fn rect_payload(rect: &RECT) -> Value {
    json!([rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top])
}

/// A caret-like rectangle at a rectangle's left edge.
fn caret_payload(edges: Edges, offset: i32) -> Value {
    json!([edges.0 + offset, edges.1, 1, edges.3 - edges.1])
}

// MARK: Pinyin in the field

fn set_attribute(shared: &Shared, context: &ITfContext, cookie: u32, range: &ITfRange) {
    let atom = shared.display_atom.get();
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
        let range = composition_range.or_else(|| selection_range(&reader, cookie));
        let rect = range.as_ref().and_then(|range| range_rect(&reader, cookie, range));
        let mut message = json!({ "t": kind, "client": client_name() });
        if let Some(rect) = rect {
            message["caret"] = rect_payload(&rect);
            message["line"] = rect_payload(&rect);
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
        let insert: ITfInsertAtSelection = writer.cast()?;
        let range = unsafe { insert.InsertTextAtSelection(cookie, TF_IAS_QUERYONLY, &[])? };
        let composer: ITfContextComposition = writer.cast()?;
        let sink = state.composition_sink.borrow().clone();
        let composition = unsafe { composer.StartComposition(cookie, &range, sink.as_ref())? };
        let range = unsafe { composition.GetRange()? };
        unsafe { range.SetText(cookie, 0, ANCHOR)? };
        select(&writer, cookie, &range)?;
        *state.formula.borrow_mut() = Some(composition);
        *state.formula_context.borrow_mut() = Some(writer.clone());
        state.last_rect.set(None);
        state.stable_rect.set(None);
        state.polls.set(0);
        unsafe {
            SetTimer(Some(state.window()), POLL_TIMER, 16, None);
        }
        log::write("formula composition started");
        Ok(())
    });
    if let Err(error) = result {
        log::write(&format!("formula composition failed: {error}"));
    }
}

/// Polls the anchor's rectangle: reported once two reads agree, then followed while the formula
/// is open (scrolling, a late layout).
fn poll_anchor(shared: &Rc<Shared>) {
    let (Some(context), Some(composition)) = (shared.formula_context.borrow().clone(), shared.formula.borrow().clone()) else {
        stop_polling(shared);
        return;
    };
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
        let Some(rect) = range_rect(&reader, cookie, &range) else { return Ok(()) };
        let current = edges(&rect);
        let previous = state.last_rect.replace(Some(current));
        if previous == Some(current) && state.stable_rect.get() != Some(current) {
            state.stable_rect.set(Some(current));
            let mut anchor = rect;
            anchor.right = anchor.left;
            state.send(json!({
                "t": "geometry",
                "client": client_name(),
                "caret": rect_payload(&anchor),
                "line": rect_payload(&rect),
            }));
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
        if let Some(reply) = reply {
            shared.send(json!({ "t": reply }));
        }
        return;
    };
    let state = shared.clone();
    let writer = context.clone();
    let result = edit::request(&context, shared.client.get(), move |cookie| {
        let range = unsafe { composition.GetRange()? };
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
    let caret = shared.stable_rect.get().map(|edges| (edges.0 + shared.formula_caret_x.get(), edges.1, edges.2, edges.3));
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
                start_formula(shared);
            } else if !on && shared.composing.get() {
                close_formula(shared, String::new(), None, None);
            }
            log::write(&format!("session {}", if on { "on" } else { "off" }));
        }
        "inline" => {
            *shared.formula_text.borrow_mut() = message.get("text").and_then(Value::as_str).unwrap_or_default().to_string();
            shared.formula_caret_x.set(message.get("caretX").and_then(Value::as_f64).unwrap_or(0.0).round() as i32);
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
        if guid.is_null() || unsafe { *guid } != display::GUID_INPUT {
            return Err(E_INVALIDARG.into());
        }
        Ok(display::AttributeInfo.into())
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
    /// keeps nothing of it), pinyin being spelled stays as it is shown, VisualTeX hears of the
    /// new field.
    fn OnSetFocus(&self, focus: Ref<'_, ITfDocumentMgr>, _previous: Ref<'_, ITfDocumentMgr>) -> Result<()> {
        let shared = &self.shared;
        if shared.composing.get() && shared.hold.get() {
            // The screenshot selection: the field gets the keyboard back after it.
            return Ok(());
        }
        if shared.composing.get() {
            close_formula(shared, String::new(), None, None);
            shared.send(json!({ "t": "forced" }));
        }
        if shared.pinyin.borrow().is_some() {
            drop_pinyin(shared, true);
        }
        shared.engine.call_existing(json!({ "op": "hide" }));
        if focus.is_some() {
            report(shared, "activate");
        } else {
            shared.send(json!({ "t": "deactivate", "client": client_name() }));
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
