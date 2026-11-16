//! The text service: VisualTeX's formula in a field. It is active in a thread only while a formula
//! is open there (host.rs activates it at VisualTeX's shortcut and gives the field its own input
//! method back when the formula ends), so it has one thing to do.
//!
//! A formula: the text service tells VisualTeX where the caret is ("activate"); VisualTeX opens the
//! formula ("session" on); the field gets a composition where the caret was, holding the formula's
//! space (space.rs: a selected hair space, which hides the field's caret, then blanks as wide as
//! each piece of the formula, joined by word spaces where the field may break the line), so the
//! field wraps the formula as its own text; the text service reads back where each piece went and
//! VisualTeX draws it there. Every key but the system's goes to VisualTeX, and "commit" replaces
//! the composition with the formula's text. However the formula ends (commit, cancel, the field
//! losing the keyboard or ending the composition itself, VisualTeX going away, nothing answering),
//! the field goes back to the input method it had.
//!
//! An application typing through IMM32 (Qt's, such as WeChat) sees the composition through
//! Windows' translation (CUAS, a "transitory" context): the anchor there is a zero-width space
//! marked as the clause being converted (the field hides its caret for it and draws nothing), and
//! the field's caret is the one the application gave IMM32 for the candidate window, where the
//! context's own rectangles are made up. Chromium's own text store (Chrome, Edge, WebView2,
//! Electron) calls itself transitory too, but it lays the composition out as its text and tells
//! where each character went: the formula wraps there as in any other field.

use crate::field::{
    add_text_height, char_rect, clear_attribute, client_name, edges, estimated_space_widths, field_extent, log_extent,
    range_extent, range_rect, rect_payload, same_line, select, selection_range, set_attribute_atom, sub_range,
    types_through_imm32, utf16, view_window, Edges, SPACE_WIDTHS,
};
use crate::link::{Link, WM_LINK_CLOSED, WM_LINK_MESSAGE};
use crate::{display, edit, host, keys, log, space};
use serde_json::{json, Value};
use std::cell::{Cell, RefCell};
use std::rc::Rc;
use windows::core::{implement, Interface, Ref, Result, BOOL, GUID};
use windows::Win32::Foundation::{E_INVALIDARG, FALSE, LPARAM, RECT, WPARAM};
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};
use windows::Win32::UI::TextServices::{
    CLSID_TF_CategoryMgr, IEnumTfDisplayAttributeInfo, ITfCategoryMgr, ITfComposition, ITfCompositionSink,
    ITfCompositionSink_Impl, ITfContext, ITfContextComposition, ITfDisplayAttributeInfo, ITfDisplayAttributeProvider,
    ITfDisplayAttributeProvider_Impl, ITfDocumentMgr, ITfInsertAtSelection, ITfKeyEventSink, ITfKeyEventSink_Impl,
    ITfKeystrokeMgr, ITfRange, ITfSource, ITfTextInputProcessor, ITfTextInputProcessorEx, ITfTextInputProcessorEx_Impl,
    ITfTextInputProcessor_Impl, ITfThreadMgr, ITfThreadMgrEventSink, ITfThreadMgrEventSink_Impl,
    INSERT_TEXT_AT_SELECTION_FLAGS, TF_ANCHOR_END, TF_IAS_QUERYONLY,
};
use windows::Win32::UI::WindowsAndMessaging::{KillTimer, SetTimer, WM_TIMER};

/// The formula's place in a field typing through IMM32: one zero-width space (the field draws the
/// clause being converted highlighted, Qt in its text colour: the anchor must have no width).
const IMM_ANCHOR: &[u16] = &[0x200B];
/// The formula's anchor is followed (Chrome lays the composition out a frame or so later; the
/// field may scroll): fast at first, then slowly while the formula is open.
const POLL_TIMER: usize = 1;
const FAST_POLLS: u32 = 30;
/// VisualTeX did not open the formula it asked for (it went away, the request was stale).
const OPEN_TIMER: usize = 5;
const OPEN_WAIT_MS: u32 = 2500;
/// A field typing through IMM32 (the standard edit controls) ends whatever is being composed when
/// its thread's input method changes, and it hears of the change a moment after the text service is
/// up: a formula's composition it ends within this of the formula's opening, with no key or click
/// of the user's, is started again a moment later.
const RESTART_TIMER: usize = 6;
const RESTART_DELAY_MS: u32 = 30;
const RESTART_WITHIN: std::time::Duration = std::time::Duration::from_millis(700);
const MAX_RESTARTS: u32 = 3;

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
enum Phase {
    /// No formula: the field is on its way back to its own input method.
    #[default]
    Idle,
    /// VisualTeX was told the caret and is opening the formula.
    Opening,
    /// The field holds the formula's space.
    Composing,
}

#[derive(Default)]
struct Shared {
    thread_mgr: RefCell<Option<ITfThreadMgr>>,
    client: Cell<u32>,
    thread_sink_cookie: Cell<u32>,
    link: RefCell<Option<Link>>,
    key_sink: RefCell<Option<ITfKeyEventSink>>,
    composition_sink: RefCell<Option<ITfCompositionSink>>,
    /// The display attribute of the formula's anchor in a field typing through IMM32 (0: none).
    anchor_atom: Cell<u32>,
    /// The display attribute of the space a field holds for a formula (0: none).
    space_atom: Cell<u32>,

    phase: Cell<Phase>,
    /// The formula VisualTeX asked for (its messages carry it).
    generation: Cell<u64>,
    /// When this formula opened, and how often its composition was started again (RESTART_TIMER).
    opened_at: Cell<Option<std::time::Instant>>,
    restarts: Cell<u32>,
    formula: RefCell<Option<ITfComposition>>,
    formula_context: RefCell<Option<ITfContext>>,
    /// The formula's field types through IMM32 (a transitory context).
    formula_transitory: Cell<bool>,
    /// What the formula inserts, for a field that ends the composition itself.
    formula_text: RefCell<String>,
    /// The formula's pieces (physical pixels wide, from VisualTeX), the composition holding them
    /// (space.rs), and the widths of a filler and of a joint as this field draws them.
    piece_widths: RefCell<Vec<f64>>,
    space: RefCell<space::Layout>,
    filler_width: Cell<f64>,
    joint_width: Cell<f64>,
    /// Calibrations in this formula: a few at most (a field rounding places to pixels must not make
    /// the space flip between two widths).
    calibrations: Cell<u32>,
    /// When the space was last held: the field draws it a moment later (follow_pieces).
    last_hold: Cell<Option<std::time::Instant>>,
    /// Where each piece started at the last poll, and what was last reported.
    last_pieces: RefCell<Vec<Edges>>,
    reported_pieces: RefCell<Option<Value>>,
    /// The widths changed at the last poll: the space is held again.
    rehold: Cell<bool>,
    last_rect: Cell<Option<Edges>>,
    polls: Cell<u32>,
    stable_rect: Cell<Option<Edges>>,
    /// VisualTeX's screenshot selection has the foreground for a moment: the formula stays open
    /// when the field loses the keyboard meanwhile.
    hold: Cell<bool>,
}

thread_local! {
    /// This thread's active text service.
    static ACTIVE: RefCell<Option<Rc<Shared>>> = const { RefCell::new(None) };
}

fn active() -> Option<Rc<Shared>> {
    ACTIVE.with(|active| active.borrow().clone())
}

impl Shared {
    /// Sends a message to VisualTeX, with the formula it is about.
    fn send(&self, mut message: Value) -> bool {
        message["gen"] = json!(self.generation.get());
        self.link.borrow().as_ref().map(|link| link.send(&message)).unwrap_or(false)
    }

    fn focused_context(&self) -> Option<ITfContext> {
        let manager = self.thread_mgr.borrow().clone()?;
        let document: ITfDocumentMgr = unsafe { manager.GetFocus() }.ok()?;
        unsafe { document.GetTop() }.ok()
    }
}

fn kill_timer(id: usize) {
    unsafe {
        let _ = KillTimer(Some(host::window()), id);
    }
}

/// The app's request for a formula in this thread, when the text service is active here already.
pub(crate) fn begin_in_this_thread(generation: u64) -> bool {
    let Some(shared) = active() else { return false };
    begin(&shared, generation);
    true
}

/// A formula opens: VisualTeX is told where the caret is and opens it there.
fn begin(shared: &Rc<Shared>, generation: u64) {
    if shared.phase.get() != Phase::Idle {
        // the shortcut again while a formula is open in this field: that one goes, without text
        finish_formula(shared, String::new(), None, false);
    }
    shared.generation.set(generation);
    shared.phase.set(Phase::Opening);
    shared.opened_at.set(Some(std::time::Instant::now()));
    shared.restarts.set(0);
    shared.hold.set(false);
    shared.formula_text.borrow_mut().clear();
    shared.piece_widths.borrow_mut().clear();
    shared.calibrations.set(0);
    *shared.space.borrow_mut() = space::Layout::default();
    if shared.link.borrow().is_none() {
        *shared.link.borrow_mut() = Link::connect(host::window());
    }
    if shared.link.borrow().is_none() {
        log::write(&format!("formula gen={generation}: VisualTeX is not running"));
        close(shared, None);
        return;
    }
    shared.send(json!({ "t": "hello", "version": env!("CARGO_PKG_VERSION") }));
    report(shared, "activate");
    unsafe {
        SetTimer(Some(host::window()), OPEN_TIMER, OPEN_WAIT_MS, None);
    }
}

/// The formula is over (its composition already ended): VisualTeX hears `tell`, and the field gets
/// its input method back.
fn close(shared: &Rc<Shared>, tell: Option<&'static str>) {
    stop_polling();
    kill_timer(OPEN_TIMER);
    kill_timer(RESTART_TIMER);
    shared.phase.set(Phase::Idle);
    if let Some(tell) = tell {
        shared.send(json!({ "t": tell, "client": client_name() }));
    }
    host::hand_back();
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
        close(shared, Some("deactivate"));
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
        // (set before the text goes in: a field that ends the composition at once finds it)
        *state.formula.borrow_mut() = Some(composition);
        *state.formula_context.borrow_mut() = Some(writer.clone());
        state.formula_transitory.set(transitory);
        let variant = std::env::var("VTX_TIP_VARIANT").unwrap_or_default();
        if !variant.is_empty() {
            log::write(&format!("DEBUG variant={variant}"));
        }
        if transitory {
            // a field typing through IMM32 reports made-up rectangles for parts of the
            // composition: the formula stays on one line there, over a zero-width anchor
            let text: &[u16] = if variant.contains("space") { &[0x200A, 0x2800] } else if variant.contains("letter") { &[0x0061] } else { IMM_ANCHOR };
            unsafe { range.SetText(cookie, 0, text)? };
            if !variant.contains("noattr") {
                set_attribute_atom(if variant.contains("spaceattr") { state.space_atom.get() } else { state.anchor_atom.get() }, &writer, cookie, &range);
            }
            if variant.contains("selend") {
                let end = unsafe { range.Clone()? };
                unsafe { end.Collapse(cookie, TF_ANCHOR_END)? };
                select(&writer, cookie, &end)?;
            } else if !variant.contains("nosel") {
                select(&writer, cookie, &range)?;
            }
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
        state.last_rect.set(None);
        state.stable_rect.set(None);
        state.last_pieces.borrow_mut().clear();
        *state.reported_pieces.borrow_mut() = None;
        state.calibrations.set(0);
        state.polls.set(0);
        unsafe {
            SetTimer(Some(host::window()), POLL_TIMER, if variant.contains("latepoll") { 600 } else { 16 }, None);
        }
        log::write(if transitory { "formula composition started (IMM32 field)" } else { "formula composition started" });
        Ok(())
    });
    if let Err(error) = result {
        log::write(&format!("formula composition failed: {error}"));
        shared.formula.borrow_mut().take();
        shared.formula_context.borrow_mut().take();
        close(shared, Some("deactivate"));
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
        state.last_hold.set(Some(std::time::Instant::now()));
        state.last_pieces.borrow_mut().clear();
        Ok(())
    });
    if let Err(error) = result {
        log::write(&format!("formula space not held: {error}"));
        return;
    }
    shared.polls.set(0);
    unsafe {
        SetTimer(Some(host::window()), POLL_TIMER, 16, None);
    }
}

/// Where the field put each piece (the first character of each): reported once two reads agree,
/// with where the held space starts and ends; the filler's and the joint's widths are read off the
/// same characters (true: the space is to be held again with them).
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
    // The widths as the field draws them, read once it has drawn the space held last: a field
    // answers some places of the composition it replaced for a moment, and widths read then are
    // wrong.
    if state.last_hold.get().map(|at| at.elapsed() < std::time::Duration::from_millis(120)).unwrap_or(false) {
        return false;
    }
    // A filler from the fillers of all the pieces read whole on their lines together (a formula of
    // short pieces holds two or three fillers each, a run long enough for a field's pixel rounding
    // to be a small part of it), a joint from the middle of those between a piece and the next on
    // its line once two of them agree.
    let mut filler_count = 0usize;
    let mut filler_run = 0.0f64;
    let mut joints: Vec<f64> = Vec::new();
    for (index, &start) in held.starts.iter().enumerate().take(rects.len()) {
        let count = held.counts[index];
        let Some(last) = char_rect(reader, cookie, range, start + count - 1) else { continue };
        if !same_line(&last, &rects[index]) || last.right <= rects[index].left {
            continue;
        }
        filler_count += count;
        filler_run += (last.right - rects[index].left) as f64;
        if index + 1 < rects.len() && same_line(&rects[index + 1], &last) && rects[index + 1].left > last.right {
            joints.push((rects[index + 1].left - last.right) as f64);
        }
    }
    let filler = (filler_count >= 6)
        .then(|| (filler_count, filler_run / filler_count as f64))
        .filter(|(_, width)| (0.5..60.0).contains(width));
    joints.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let joint = joints
        .get(joints.len() / 2)
        .copied()
        .filter(|middle| (1.0..60.0).contains(middle) && joints.iter().filter(|width| (*width - middle).abs() <= 1.5).count() >= 2);
    let window = view_window(reader);
    let filler_changed = filler.map(|(count, width)| (width - state.filler_width.get()).abs() * count as f64 > 1.5).unwrap_or(false);
    let joint_changed = joint.map(|width| (width - state.joint_width.get()).abs() > 1.5).unwrap_or(false);
    if !filler_changed && !joint_changed {
        if filler.is_some() || joint.is_some() {
            // the widths the space is held with: this window's next formulas start with them
            SPACE_WIDTHS.with(|widths| widths.borrow_mut().insert(window, (state.filler_width.get(), state.joint_width.get())));
        }
        return false;
    }
    // past three, the widths stay those the space is held with (never a space held with other
    // widths than the ones the pieces are placed with)
    if state.calibrations.get() >= 3 {
        return false;
    }
    state.calibrations.set(state.calibrations.get() + 1);
    if let Some((count, width)) = filler.filter(|_| filler_changed) {
        log::write(&format!("filler {:.2} -> {width:.2} from {count}", state.filler_width.get()));
        state.filler_width.set(width);
    }
    if let Some(width) = joint.filter(|_| joint_changed) {
        log::write(&format!("joint {:.2} -> {width:.2} from {}", state.joint_width.get(), joints.len()));
        state.joint_width.set(width);
    }
    SPACE_WIDTHS.with(|widths| widths.borrow_mut().insert(window, (state.filler_width.get(), state.joint_width.get())));
    true
}

/// Polls the anchor's rectangle: reported once two reads agree, then followed while the formula
/// is open (scrolling, a late layout).
fn poll_anchor(shared: &Rc<Shared>) {
    let (Some(context), Some(composition)) = (shared.formula_context.borrow().clone(), shared.formula.borrow().clone()) else {
        stop_polling();
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
            SetTimer(Some(host::window()), POLL_TIMER, 250, None);
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

fn stop_polling() {
    kill_timer(POLL_TIMER);
}

/// Ends the formula's composition with `text` in its place (empty: removed), tells VisualTeX
/// `reply`, and (`then_close`) gives the field its input method back, after the edit.
fn finish_formula(shared: &Rc<Shared>, text: String, reply: Option<&'static str>, then_close: bool) {
    stop_polling();
    let composition = shared.formula.borrow_mut().take();
    let context = shared.formula_context.borrow_mut().take();
    let (Some(composition), Some(context)) = (composition, context) else {
        // The composition never started or the field dropped it unseen: a formula committed then
        // still goes in at the field's caret.
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
        if then_close {
            close(shared, reply);
        } else if let Some(reply) = reply {
            shared.send(json!({ "t": reply }));
        }
        return;
    };
    let state = shared.clone();
    let writer = context.clone();
    let done = Rc::new(Cell::new(false));
    let ran = done.clone();
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
        ran.set(true);
        // (the input method goes back only now: handed back before this edit ran, the field would
        // end the composition itself, without the formula's text)
        if then_close {
            close(&state, reply);
        } else if let Some(reply) = reply {
            state.send(json!({ "t": reply }));
        }
        Ok(())
    });
    if let Err(error) = &result {
        log::write(&format!("formula commit failed: {error}"));
    }
    // (an edit TSF refused, or one that failed half way)
    if result.is_err() && !done.get() && then_close {
        close(shared, reply);
    }
}

/// A message from VisualTeX, on the field's thread.
fn receive(shared: &Rc<Shared>, message: &Value) {
    let Some(kind) = message.get("t").and_then(Value::as_str) else { return };
    match kind {
        "session" => {
            let on = message.get("on").and_then(Value::as_bool).unwrap_or(false);
            log::write(&format!("session {} phase={:?}", if on { "on" } else { "off" }, shared.phase.get()));
            if on && shared.phase.get() == Phase::Opening {
                kill_timer(OPEN_TIMER);
                shared.phase.set(Phase::Composing);
                start_formula(shared);
            } else if !on && shared.phase.get() != Phase::Idle {
                finish_formula(shared, String::new(), None, true);
            }
        }
        "inline" => {
            *shared.formula_text.borrow_mut() = message.get("text").and_then(Value::as_str).unwrap_or_default().to_string();
            if let Some(pieces) = message.get("pieces").and_then(Value::as_array) {
                let widths: Vec<f64> = pieces.iter().map(|piece| piece.as_f64().unwrap_or(0.0).clamp(0.0, 16384.0)).collect();
                if shared.phase.get() == Phase::Composing && *shared.piece_widths.borrow() != widths {
                    *shared.piece_widths.borrow_mut() = widths;
                    hold(shared);
                }
            }
        }
        "commit" => {
            let text = message.get("text").and_then(Value::as_str).unwrap_or_default().to_string();
            log::write(&format!("commit {} chars", text.chars().count()));
            finish_formula(shared, text, Some("committed"), true);
        }
        "cancel" => {
            log::write("cancel");
            finish_formula(shared, String::new(), Some("cancelled"), true);
        }
        "caret" => report(shared, "geometry"),
        "hold" => shared.hold.set(message.get("on").and_then(Value::as_bool).unwrap_or(false)),
        _ => {}
    }
}

/// A message of this thread's window that is the text service's (true: handled).
pub(crate) fn window_message(message: u32, wparam: WPARAM, lparam: LPARAM) -> bool {
    match message {
        WM_LINK_MESSAGE => {
            let text = unsafe { Box::from_raw(lparam.0 as *mut String) };
            if let (Some(shared), Ok(value)) = (active(), serde_json::from_str::<Value>(&text)) {
                receive(&shared, &value);
            }
            true
        }
        WM_LINK_CLOSED => {
            if let Some(shared) = active() {
                shared.link.borrow_mut().take();
                log::write("link closed");
                if shared.phase.get() != Phase::Idle {
                    finish_formula(&shared, String::new(), None, true);
                }
            }
            true
        }
        WM_TIMER => {
            let Some(shared) = active() else {
                kill_timer(wparam.0);
                return true;
            };
            match wparam.0 {
                POLL_TIMER => poll_anchor(&shared),
                OPEN_TIMER => {
                    kill_timer(OPEN_TIMER);
                    if shared.phase.get() == Phase::Opening {
                        log::write("VisualTeX did not open the formula");
                        close(&shared, None);
                    }
                }
                RESTART_TIMER => {
                    kill_timer(RESTART_TIMER);
                    if shared.phase.get() == Phase::Composing && shared.formula.borrow().is_none() {
                        *shared.space.borrow_mut() = space::Layout::default();
                        start_formula(&shared);
                        // (the pieces VisualTeX told meanwhile)
                        hold(&shared);
                    }
                }
                _ => return false,
            }
            true
        }
        _ => false,
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

    /// The field's thread went to another input method: the hand-back after a formula, or the user
    /// choosing one while a formula was open (the formula goes, and VisualTeX's profile is disabled
    /// again all the same).
    fn Deactivate(&self) -> Result<()> {
        let shared = self.shared.borrow().clone();
        if shared.phase.get() != Phase::Idle {
            log::write("deactivated with a formula open");
            finish_formula(&shared, String::new(), None, false);
            stop_polling();
            kill_timer(OPEN_TIMER);
            shared.phase.set(Phase::Idle);
            shared.send(json!({ "t": "deactivate", "client": client_name() }));
        }
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
        shared.key_sink.borrow_mut().take();
        shared.composition_sink.borrow_mut().take();
        shared.link.borrow_mut().take();
        ACTIVE.with(|active| {
            let mut active = active.borrow_mut();
            if active.as_ref().map(|current| Rc::ptr_eq(current, &shared)).unwrap_or(false) {
                *active = None;
            }
        });
        // A fresh state for a later activation of this object.
        *self.shared.borrow_mut() = Rc::new(Shared::default());
        host::hand_back();
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
        host::ensure();
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
            if let Ok(atom) = unsafe { categories.RegisterGUID(&display::GUID_ANCHOR) } {
                shared.anchor_atom.set(atom);
            }
            if let Ok(atom) = unsafe { categories.RegisterGUID(&display::GUID_SPACE) } {
                shared.space_atom.set(atom);
            }
        }
        ACTIVE.with(|active| *active.borrow_mut() = Some(shared.clone()));
        let generation = host::take_pending();
        log::write(&format!("activated gen={generation}"));
        if generation == 0 {
            // Not VisualTeX's shortcut (the profile chosen in the input switcher while a formula was
            // open somewhere): VisualTeX types formulas only, the thread goes back at once.
            host::hand_back();
        } else {
            begin(&shared, generation);
        }
        Ok(())
    }
}

#[implement(ITfKeyEventSink)]
struct KeySink {
    shared: Rc<Shared>,
}

impl KeySink {
    /// While a formula is open (or opening: VisualTeX keeps those keys for it) its keys are
    /// VisualTeX's, but the system's and the application's Control chords, and the modifier keys
    /// themselves (Alt alone is taken: it would put the keyboard in the application's menu).
    fn takes(&self, vk: u32) -> bool {
        if self.shared.phase.get() == Phase::Idle || keys::belongs_to_system(vk) {
            return false;
        }
        match vk {
            0x10 | 0xa0 | 0xa1 | 0x11 | 0xa2 | 0xa3 | 0x14 | 0x5b | 0x5c => false,
            0x12 | 0xa4 | 0xa5 => !keys::modifiers().ctrl,
            _ => keys::belongs_to_formula(vk),
        }
    }

    fn key(&self, vk: u32, lparam: isize, release: bool) -> bool {
        if !self.takes(vk) {
            return false;
        }
        if !release && !matches!(vk, 0x12 | 0xa4 | 0xa5) {
            let repeat = (lparam >> 30) & 1 == 1;
            self.shared.send(keys::message(vk, lparam, repeat));
        }
        true
    }
}

impl ITfKeyEventSink_Impl for KeySink_Impl {
    fn OnSetFocus(&self, _foreground: BOOL) -> Result<()> {
        Ok(())
    }

    fn OnTestKeyDown(&self, _context: Ref<'_, ITfContext>, wparam: WPARAM, _lparam: LPARAM) -> Result<BOOL> {
        Ok(BOOL::from(self.takes(wparam.0 as u32)))
    }

    fn OnTestKeyUp(&self, _context: Ref<'_, ITfContext>, wparam: WPARAM, _lparam: LPARAM) -> Result<BOOL> {
        Ok(BOOL::from(self.takes(wparam.0 as u32)))
    }

    fn OnKeyDown(&self, _context: Ref<'_, ITfContext>, wparam: WPARAM, lparam: LPARAM) -> Result<BOOL> {
        Ok(BOOL::from(self.key(wparam.0 as u32, lparam.0, false)))
    }

    fn OnKeyUp(&self, _context: Ref<'_, ITfContext>, wparam: WPARAM, lparam: LPARAM) -> Result<BOOL> {
        Ok(BOOL::from(self.key(wparam.0 as u32, lparam.0, true)))
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

    /// The keyboard went to another field (or away): the formula closes (the field it belonged to
    /// keeps nothing of it; VisualTeX keeps the draft for the next formula) and the thread goes
    /// back to its input method.
    fn OnSetFocus(&self, focus: Ref<'_, ITfDocumentMgr>, previous: Ref<'_, ITfDocumentMgr>) -> Result<()> {
        let shared = &self.shared;
        if shared.phase.get() == Phase::Idle {
            return Ok(());
        }
        // (the same document again is no change: some fields set their focus anew when an input
        // method comes up)
        let same = match (focus.as_ref(), previous.as_ref()) {
            (Some(focus), Some(previous)) => focus.as_raw() == previous.as_raw(),
            _ => false,
        };
        log::write(&format!("focus changed: field={} same={same} hold={} phase={:?}", focus.is_some(), shared.hold.get(), shared.phase.get()));
        if same || shared.hold.get() {
            // (the screenshot selection: the field gets the keyboard back after it)
            return Ok(());
        }
        finish_formula(shared, String::new(), Some("deactivate"), true);
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
    /// The field ended the composition itself (a click elsewhere in it): it gets the formula's
    /// text, not the held space.
    fn OnCompositionTerminated(&self, cookie: u32, composition: Ref<'_, ITfComposition>) -> Result<()> {
        let shared = &self.shared;
        let is_formula = match (composition.as_ref(), shared.formula.borrow().as_ref()) {
            (Some(ended), Some(formula)) => ended.as_raw() == formula.as_raw(),
            _ => false,
        };
        if !is_formula || shared.phase.get() == Phase::Idle {
            return Ok(());
        }
        shared.formula.borrow_mut().take();
        shared.formula_context.borrow_mut().take();
        // The field heard of its thread's new input method and dropped what was being composed (see
        // RESTART_TIMER): the formula's space is taken out and held again in a moment.
        let just_opened = shared.opened_at.get().map(|at| at.elapsed() < RESTART_WITHIN).unwrap_or(false);
        if just_opened && shared.formula_transitory.get() && shared.restarts.get() < MAX_RESTARTS {
            shared.restarts.set(shared.restarts.get() + 1);
            stop_polling();
            if let Some(range) = composition.as_ref().and_then(|composition| unsafe { composition.GetRange() }.ok()) {
                if let Ok(context) = unsafe { range.GetContext() } {
                    clear_attribute(&context, cookie, &range);
                }
                unsafe {
                    let _ = range.SetText(cookie, 0, &[]);
                }
            }
            log::write(&format!("the field ended the composition as the formula opened: started again ({})", shared.restarts.get()));
            unsafe {
                SetTimer(Some(host::window()), RESTART_TIMER, RESTART_DELAY_MS, None);
            }
            return Ok(());
        }
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
        log::write(&format!("formula composition terminated by the field (polls={})", shared.polls.get()));
        close(shared, Some("forced"));
        Ok(())
    }
}
