//! The composition strip's window: a borderless, transparent, non-activating
//! panel created once and kept hydrated, like the resident Office editors
//! (office/macos_offline.rs). Parked it stays ordered in behind everything at
//! a near-zero alpha and ignores the mouse, so WebKit keeps the page running
//! and the next activation only moves it.

use super::{PanelLayout, ScreenRect, FRAME_EVENT, PANEL_TITLE, WINDOW_LABEL};
use objc2_app_kit::{NSScreen, NSWindow, NSWindowCollectionBehavior};
use objc2_foundation::{MainThreadMarker, NSPoint, NSRect, NSSize};
use std::sync::mpsc;
use std::time::Duration;
use tauri::utils::config::BackgroundThrottlingPolicy;
use tauri::{AppHandle, Emitter, Manager, WebviewUrl, WebviewWindow, WebviewWindowBuilder, WindowEvent};

/// Above other applications' floating panels and full-screen windows, like
/// an input method's candidate window.
const PRESENTED_LEVEL: isize = 101; // NSPopUpMenuWindowLevel
const PARKED_ALPHA: f64 = 0.001;
/// The card's distance under the caret line (input method).
const BELOW_LINE_GAP: f64 = 4.0;

/// Where the strip stands. OnLine: on the caret line, its formula starting
/// at the caret (pasting). BelowLine: under the line, its top edge (the
/// layout's anchor_y) under the caret. InLine: the formula drawn in the
/// field's line itself (input method): the anchor is the formula's start on
/// the line's baseline, and the layout's anchor is the formula's origin on
/// its baseline.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Placement {
    OnLine,
    BelowLine,
    InLine,
}
/// While the page first mounts it needs a little more than PARKED_ALPHA, or
/// WebKit suspends it (see wake_resident_editor_invisibly).
const HYDRATING_ALPHA: f64 = 0.01;

pub(super) fn ensure(app: &AppHandle) -> Result<WebviewWindow, String> {
    if let Some(window) = app.get_webview_window(WINDOW_LABEL) {
        return Ok(window);
    }
    let theme = crate::persisted_app_theme(app);
    let layout = PanelLayout::EMPTY;
    let window = WebviewWindowBuilder::new(
        app,
        WINDOW_LABEL,
        WebviewUrl::App(format!("anywhere.html?transport=tauri&theme={theme}").into()),
    )
    .title(PANEL_TITLE)
    .inner_size(layout.width, layout.height)
    .decorations(false)
    .transparent(true)
    .shadow(false)
    .resizable(false)
    .focused(false)
    .skip_taskbar(true)
    .visible(false)
    .background_throttling(BackgroundThrottlingPolicy::Disabled)
    .build()
    .map_err(|error| format!("Unable to create the VisualTeX Anywhere panel: {error}"))?;
    let events = app.clone();
    window.on_window_event(move |event| {
        if let WindowEvent::Focused(false) = event {
            super::handle_focus_lost(&events);
        }
    });
    with_native_window(&window, |native_window| {
        let behavior = native_window.collectionBehavior()
            | NSWindowCollectionBehavior::CanJoinAllSpaces
            | NSWindowCollectionBehavior::FullScreenAuxiliary
            | NSWindowCollectionBehavior::Transient
            | NSWindowCollectionBehavior::IgnoresCycle;
        native_window.setCollectionBehavior(behavior);
        native_window.setHidesOnDeactivate(false);
        native_window.setExcludedFromWindowsMenu(true);
        native_window.setHasShadow(false);
        set_parked(native_window, HYDRATING_ALPHA);
    })?;
    Ok(window)
}

fn with_native_window(
    window: &WebviewWindow,
    work: impl FnOnce(&NSWindow) + Send + 'static,
) -> Result<(), String> {
    window
        .with_webview(move |webview| unsafe {
            let native_window: &NSWindow = &*webview.ns_window().cast();
            work(native_window);
        })
        .map_err(|error| format!("Unable to reach the VisualTeX Anywhere panel: {error}"))
}

/// The panel's window number (what is beneath it is the host's, see
/// screen_colors).
pub(super) fn window_number(window: &WebviewWindow) -> Option<u32> {
    let (sender, receiver) = mpsc::sync_channel(1);
    with_native_window(window, move |native_window| {
        let _ = sender.send(native_window.windowNumber());
    })
    .ok()?;
    let number = receiver.recv_timeout(Duration::from_millis(500)).ok()?;
    u32::try_from(number).ok().filter(|number| *number > 0)
}

/// Mouse-inert, behind everything, almost transparent: hydrated but unseen.
/// AppKit's orderBack does not resign a key panel, so it is ordered out first;
/// the panel then cannot become key again until it is presented
/// (vendor/tao-office-panel: a panel that ignores the mouse is never key).
fn set_parked(native_window: &NSWindow, alpha: f64) {
    native_window.setIgnoresMouseEvents(true);
    if native_window.isKeyWindow() {
        native_window.orderOut(None);
    }
    native_window.setLevel(objc2_app_kit::NSNormalWindowLevel);
    native_window.setAlphaValue(alpha);
    native_window.orderBack(None);
}

pub(super) fn park(window: &WebviewWindow) {
    if let Err(error) = with_native_window(window, |native_window| set_parked(native_window, PARKED_ALPHA)) {
        eprintln!("{error}");
    }
}

/// Parked while its page loads again: WebKit suspends a mounting page at
/// PARKED_ALPHA.
pub(super) fn park_for_reload(window: &WebviewWindow) {
    if let Err(error) = with_native_window(window, |native_window| set_parked(native_window, HYDRATING_ALPHA)) {
        eprintln!("{error}");
    }
}

/// Parks the panel and returns once AppKit has given the keyboard back.
pub(super) fn park_and_wait(window: &WebviewWindow) -> Result<(), String> {
    let (sender, receiver) = mpsc::sync_channel(1);
    with_native_window(window, move |native_window| {
        set_parked(native_window, PARKED_ALPHA);
        let _ = sender.send(());
    })?;
    receiver
        .recv_timeout(Duration::from_millis(500))
        .map_err(|_| "Timed out hiding the VisualTeX Anywhere panel".to_string())
}

pub(super) fn present(window: &WebviewWindow, anchor: ScreenRect, layout: PanelLayout) {
    let emitter = window.clone();
    let presented = window.with_webview(move |webview| unsafe {
        let native_window: &NSWindow = &*webview.ns_window().cast();
        let (frame, report) = frame_for(anchor, layout, Placement::OnLine);
        native_window.setFrame_display(frame, true);
        tell_page(emitter, report);
        native_window.setIgnoresMouseEvents(false);
        native_window.setAlphaValue(1.0);
        native_window.setLevel(PRESENTED_LEVEL);
        native_window.orderFrontRegardless();
        native_window.makeKeyWindow();
        let webview_responder: &objc2_app_kit::NSResponder = &*webview.inner().cast();
        native_window.makeFirstResponder(Some(webview_responder));
    });
    if let Err(error) = presented {
        eprintln!("Unable to show the VisualTeX Anywhere panel: {error}");
    }
}

/// The input method's formula: shown without taking the keyboard (the field
/// keeps it; the input method hands the keys over) and without the mouse.
pub(super) fn present_passive(window: &WebviewWindow, anchor: ScreenRect, layout: PanelLayout, placement: Placement) {
    let emitter = window.clone();
    let presented = with_native_window(window, move |native_window| {
        let (frame, report) = frame_for(anchor, layout, placement);
        native_window.setFrame_display(frame, true);
        tell_page(emitter, report);
        native_window.setIgnoresMouseEvents(true);
        native_window.setAlphaValue(1.0);
        native_window.setLevel(PRESENTED_LEVEL);
        native_window.orderFrontRegardless();
    });
    if let Err(error) = presented {
        eprintln!("Unable to show the VisualTeX Anywhere card: {error}");
    }
}

/// The strip changed size (the formula grew, candidates opened) or the caret
/// moved: keep it at its place.
pub(super) fn reframe(window: &WebviewWindow, anchor: ScreenRect, layout: PanelLayout, placement: Placement) {
    let emitter = window.clone();
    let reframed = with_native_window(window, move |native_window| {
        let (frame, report) = frame_for(anchor, layout, placement);
        native_window.setFrame_display(frame, true);
        tell_page(emitter, report);
    });
    if let Err(error) = reframed {
        eprintln!("{error}");
    }
}

/// Where the panel stands, for the page, in points from the main display's top left (as a window's screenX and
/// screenY would tell): WebKit gives a panel's window.screenX/Y as 0 and the screen's height, so the page cannot
/// tell where it is on the screen (and put every candidate list above the line). `visible_*` is the visible area of
/// the caret's display (menu bar and Dock left out).
#[derive(Clone, Copy, Debug, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct FrameReport {
    x: f64,
    y: f64,
    width: f64,
    height: f64,
    visible_left: f64,
    visible_top: f64,
    visible_right: f64,
    visible_bottom: f64,
}

/// Tells the page where the panel stands. From a thread of its own, never inside the main thread's closure that sets the
/// frame: emitting evaluates script on the webview through a lock that closure's own dispatch holds, so the main thread
/// waited for itself (the app hung at the first card).
fn tell_page(window: WebviewWindow, report: FrameReport) {
    let _ = std::thread::Builder::new()
        .name("visualtex-anywhere-frame".to_string())
        .spawn(move || {
            let _ = window.emit(FRAME_EVENT, report);
        });
}

/// The panel frame (AppKit coordinates: origin at the bottom left of the main
/// display, y up) that puts the strip's anchor on the caret's vertical
/// centre (OnLine) or its top edge under the caret line (BelowLine), kept on
/// the caret's display, and where that is, for the page.
fn frame_for(anchor: ScreenRect, layout: PanelLayout, placement: Placement) -> (NSRect, FrameReport) {
    let main_height = main_display_height();
    let left = anchor.x - layout.anchor_x;
    let top = match placement {
        Placement::OnLine => anchor.y + anchor.height / 2.0 - layout.anchor_y,
        Placement::BelowLine => anchor.y + anchor.height + BELOW_LINE_GAP - layout.anchor_y,
        Placement::InLine => anchor.y - layout.anchor_y,
    };
    let mut frame = NSRect::new(
        NSPoint::new(left, main_height - top - layout.height),
        NSSize::new(layout.width, layout.height),
    );
    let caret_point = NSPoint::new(anchor.x, main_height - anchor.y - anchor.height / 2.0);
    let visible_area = visible_frame_containing(caret_point);
    if let Some(visible) = visible_area {
        if placement != Placement::InLine {
            let max_x = visible.origin.x + visible.size.width - frame.size.width;
            frame.origin.x = frame.origin.x.min(max_x).max(visible.origin.x);
        }
        if placement == Placement::BelowLine && frame.origin.y < visible.origin.y {
            // No room under the line: the card goes over it, as candidate
            // windows do at the bottom of the screen.
            frame.origin.y = main_height - anchor.y + BELOW_LINE_GAP;
        }
        // Too low for the strip and its candidates: rise above the screen's
        // bottom edge (the caret line stays visible above the candidates).
        // A formula drawn in the line stays on it.
        if placement != Placement::InLine {
            let max_y = visible.origin.y + visible.size.height - frame.size.height;
            frame.origin.y = frame.origin.y.max(visible.origin.y).min(max_y.max(visible.origin.y));
        }
    }
    // no display found (cannot happen with a caret on a screen): a visible area that holds everything
    const UNBOUNDED: f64 = 100_000.0;
    let (visible_left, visible_top, visible_right, visible_bottom) = match visible_area {
        Some(visible) => (
            visible.origin.x,
            main_height - (visible.origin.y + visible.size.height),
            visible.origin.x + visible.size.width,
            main_height - visible.origin.y,
        ),
        None => (-UNBOUNDED, -UNBOUNDED, UNBOUNDED, UNBOUNDED),
    };
    let report = FrameReport {
        x: frame.origin.x,
        y: main_height - (frame.origin.y + frame.size.height),
        width: frame.size.width,
        height: frame.size.height,
        visible_left,
        visible_top,
        visible_right,
        visible_bottom,
    };
    (frame, report)
}

pub(super) fn main_display_height() -> f64 {
    #[repr(C)]
    struct CGRect {
        x: f64,
        y: f64,
        width: f64,
        height: f64,
    }
    #[link(name = "CoreGraphics", kind = "framework")]
    unsafe extern "C" {
        fn CGMainDisplayID() -> u32;
        fn CGDisplayBounds(display: u32) -> CGRect;
    }
    unsafe { CGDisplayBounds(CGMainDisplayID()).height }
}

fn visible_frame_containing(point: NSPoint) -> Option<NSRect> {
    let main_thread = MainThreadMarker::new()?;
    let screens = NSScreen::screens(main_thread);
    let mut nearest: Option<(f64, NSRect)> = None;
    for index in 0..screens.count() {
        let screen = screens.objectAtIndex(index);
        let frame = screen.frame();
        let dx = (frame.origin.x - point.x).max(point.x - (frame.origin.x + frame.size.width)).max(0.0);
        let dy = (frame.origin.y - point.y).max(point.y - (frame.origin.y + frame.size.height)).max(0.0);
        let distance = dx * dx + dy * dy;
        let closer = match nearest {
            Some((best, _)) => distance < best,
            None => true,
        };
        if closer {
            nearest = Some((distance, screen.visibleFrame()));
        }
    }
    nearest.map(|(_, visible)| visible)
}
