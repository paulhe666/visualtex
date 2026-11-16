//! What a field tells about its text: where a range of it is on the screen, in physical pixels, for
//! every kind of field (TSF's own, Chromium's, and those typing through IMM32, whose rectangles
//! Windows makes up and whose real caret is the one the application gave IMM32), the height of a
//! line of its text, and the widths it draws the formula's blanks with.

use crate::log;
use serde_json::{json, Value};
use std::cell::RefCell;
use std::collections::HashMap;
use std::mem::ManuallyDrop;
use windows::core::{Result, BOOL};
use windows::Win32::Foundation::{FALSE, HWND, POINT, RECT};
use windows::Win32::Graphics::Gdi::ClientToScreen;
use windows::Win32::System::Variant::{VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0, VT_I4};
use windows::Win32::UI::HiDpi::LogicalToPhysicalPointForPerMonitorDPI;
use windows::Win32::UI::Input::Ime::{ImmGetCandidateWindow, ImmGetContext, ImmReleaseContext, CANDIDATEFORM, CFS_EXCLUDE};
use windows::Win32::UI::Input::KeyboardAndMouse::GetFocus;
use windows::Win32::UI::TextServices::{
    ITfContext, ITfRange, GUID_PROP_ATTRIBUTE, TF_AE_NONE, TF_ANCHOR_START, TF_DEFAULT_SELECTION, TF_SELECTION,
    TF_SELECTIONSTYLE, TF_SS_TRANSITORY,
};
use windows::Win32::UI::WindowsAndMessaging::GetClassNameW;

pub(crate) type Edges = (i32, i32, i32, i32);

pub(crate) fn client_name() -> String {
    std::env::current_exe()
        .ok()
        .and_then(|exe| exe.file_stem().map(|name| name.to_string_lossy().into_owned()))
        .unwrap_or_default()
}

pub(crate) fn utf16(text: &str) -> Vec<u16> {
    text.encode_utf16().collect()
}

/// Where a range of the field's text is on the screen.
pub(crate) struct Extent {
    /// Physical pixels (top-left origin), whatever the application's DPI awareness.
    pub(crate) rect: RECT,
    /// The field cut the rectangle to what it shows (a line partly scrolled out of view is shorter
    /// than its text).
    pub(crate) clipped: bool,
    /// The view's window (0: none).
    pub(crate) window: isize,
}

/// The screen rectangle of `range` (see Extent).
pub(crate) fn range_extent(context: &ITfContext, cookie: u32, range: &ITfRange) -> Option<Extent> {
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
pub(crate) fn range_rect(context: &ITfContext, cookie: u32, range: &ITfRange) -> Option<RECT> {
    range_extent(context, cookie, range).map(|extent| extent.rect)
}

/// The context is the one Windows keeps for an application typing through IMM32 (CUAS): it holds
/// the composition only, and its rectangles are made up from the application's composition window
/// (one line of the default font below the caret where the application gives no font, as Qt).
pub(crate) fn is_transitory(context: &ITfContext) -> bool {
    unsafe { context.GetStatus() }.map(|status| status.dwStaticFlags & TF_SS_TRANSITORY != 0).unwrap_or(false)
}

/// The window of the context's view (0: none).
pub(crate) fn view_window(context: &ITfContext) -> isize {
    unsafe { context.GetActiveView().and_then(|view| view.GetWnd()) }
        .ok()
        .filter(|window| !window.is_invalid())
        .map(|window| window.0 as isize)
        .unwrap_or(0)
}

/// The view's window is Chromium's (Chrome_WidgetWin_1, Chrome_RenderWidgetHostHWND: Chrome, Edge,
/// WebView2, Electron applications).
pub(crate) fn chromium_view(context: &ITfContext) -> bool {
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
pub(crate) fn types_through_imm32(context: &ITfContext) -> bool {
    is_transitory(context) && !chromium_view(context)
}

/// The caret an application typing through IMM32 gave Windows for the candidate window: the area
/// the candidates must not cover (Qt sets its text cursor's rectangle there), in physical screen
/// pixels. Only a caret-like area counts (some applications exclude their whole text area). Read in
/// the application's own UI thread, as IMM32 requires.
pub(crate) fn imm_caret(window: isize) -> Option<RECT> {
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
pub(crate) fn made_up(rect: &RECT, caret: &RECT) -> bool {
    let caret_height = caret.bottom - caret.top;
    let height = rect.bottom - rect.top;
    let overlap = rect.bottom.min(caret.bottom) - rect.top.max(caret.top);
    let near = (rect.left - caret.left).abs() <= caret_height * 2 && (rect.top - caret.top).abs() <= caret_height * 2;
    near && (overlap * 2 < caret_height || height * 4 < caret_height * 3 || height * 3 > caret_height * 4)
}

/// The extent to report for a field: `extent` (the context's), or in a transitory context the
/// IMM32 caret where that is made up. Some(note): the caret is IMM32's (the note says what it
/// replaced, for the log).
pub(crate) fn field_extent(context: &ITfContext, extent: Option<Extent>) -> (Option<Extent>, Option<String>) {
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

pub(crate) fn edges(rect: &RECT) -> Edges {
    (rect.left, rect.top, rect.right, rect.bottom)
}

pub(crate) fn selection_range(context: &ITfContext, cookie: u32) -> Option<ITfRange> {
    let mut selection = [TF_SELECTION::default()];
    let mut fetched = 0u32;
    unsafe { context.GetSelection(cookie, TF_DEFAULT_SELECTION, &mut selection, &mut fetched) }.ok()?;
    if fetched == 0 {
        return None;
    }
    unsafe { ManuallyDrop::take(&mut selection[0].range) }
}

pub(crate) fn select(context: &ITfContext, cookie: u32, range: &ITfRange) -> Result<()> {
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
    pub(crate) static SPACE_WIDTHS: RefCell<HashMap<isize, (f64, f64)>> = RefCell::new(HashMap::new());
    /// The height of a line of text by window, learnt from the caret where the line had text
    /// before it (physical pixels). An empty line reports its whole line box instead (Chrome,
    /// Electron), which would make the formula too large: VisualTeX takes this height there (as
    /// the macOS input method does). Kept across activations (the service is made anew each
    /// time); by window, so that one field's text (a search box's) does not size another's.
    pub(crate) static TEXT_HEIGHTS: RefCell<HashMap<isize, i32>> = RefCell::new(HashMap::new());
}

/// Text stands before `range` on its line.
pub(crate) fn text_before(cookie: u32, range: &ITfRange) -> bool {
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
pub(crate) fn add_text_height(message: &mut Value, cookie: u32, range: &ITfRange, extent: &Extent, learn: bool) {
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
pub(crate) fn log_extent(what: &str, extent: &Extent, message: &Value, source: Option<&str>) {
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

pub(crate) fn rect_payload(rect: &RECT) -> Value {
    json!([rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top])
}

/// `length` characters of `range` from its `start`-th.
pub(crate) fn sub_range(range: &ITfRange, cookie: u32, start: usize, length: usize) -> Option<ITfRange> {
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
pub(crate) fn char_rect(context: &ITfContext, cookie: u32, range: &ITfRange, index: usize) -> Option<RECT> {
    let part = sub_range(range, cookie, index, 1)?;
    range_rect(context, cookie, &part).filter(|rect| rect.right >= rect.left && rect.bottom - rect.top < 2000)
}

/// Two rectangles on one line of the field.
pub(crate) fn same_line(a: &RECT, b: &RECT) -> bool {
    (a.bottom - b.bottom).abs() * 2 < (a.bottom - a.top).max(1)
}

/// A filler's and a joint's width before the field has shown them: from the height of a line of
/// its text (a blank is about half an em, a word space a quarter).
pub(crate) fn estimated_space_widths(window: isize, line_height: i32) -> (f64, f64) {
    if let Some(known) = SPACE_WIDTHS.with(|widths| widths.borrow().get(&window).copied()) {
        return known;
    }
    let height = TEXT_HEIGHTS.with(|heights| heights.borrow().get(&window).copied()).unwrap_or(line_height).clamp(8, 200) as f64;
    let em = height / 1.32;
    (em * 0.5, em * 0.27)
}

/// Gives `range` the display attribute registered as `atom` (0: none).
pub(crate) fn set_attribute_atom(atom: u32, context: &ITfContext, cookie: u32, range: &ITfRange) {
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

pub(crate) fn clear_attribute(context: &ITfContext, cookie: u32, range: &ITfRange) {
    if let Ok(property) = unsafe { context.GetProperty(&GUID_PROP_ATTRIBUTE) } {
        unsafe {
            let _ = property.Clear(cookie, range);
        }
    }
}
