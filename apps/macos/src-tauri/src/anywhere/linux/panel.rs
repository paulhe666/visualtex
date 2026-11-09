//! The composition strip's window on Linux (X11, also under XWayland): borderless, transparent,
//! above other windows, created once and kept loaded. The input method's formula is shown without
//! taking the keyboard and lets the mouse through (the field keeps both); the pasting strip takes the
//! keyboard. Places are screen pixels, y down, as the input method reports them.

use super::{PanelLayout, ScreenRect, PANEL_TITLE, WINDOW_LABEL};
use gtk::prelude::*;
use std::sync::OnceLock;
use std::time::Duration;
use tauri::utils::config::BackgroundThrottlingPolicy;
use tauri::{AppHandle, Manager, WebviewUrl, WebviewWindow, WebviewWindowBuilder, WindowEvent};

/// The card's distance under the caret line (input method).
const BELOW_LINE_GAP: f64 = 4.0;

static APP: OnceLock<AppHandle> = OnceLock::new();

pub(super) fn app() -> Option<&'static AppHandle> {
    APP.get()
}

/// Where the strip stands (as on macOS, panel.rs there).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Placement {
    OnLine,
    BelowLine,
    InLine,
}

pub(super) fn ensure(app: &AppHandle) -> Result<WebviewWindow, String> {
    let _ = APP.set(app.clone());
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
    .always_on_top(true)
    .skip_taskbar(true)
    .visible(false)
    .background_throttling(BackgroundThrottlingPolicy::Disabled)
    .build()
    .map_err(|error| format!("Unable to create the VisualTeX Anywhere panel: {error}"))?;
    let events = app.clone();
    window.on_window_event(move |event| {
        if let WindowEvent::Focused(false) = event {
            // only the pasting strip takes the keyboard: losing it hides that strip
            if TOOK_KEYBOARD.swap(false, std::sync::atomic::Ordering::SeqCst) {
                super::handle_focus_lost(&events);
            }
        }
    });
    with_gtk(&window, |gtk_window| {
        gtk_window.set_type_hint(gdk::WindowTypeHint::Utility);
        gtk_window.set_skip_pager_hint(true);
        gtk_window.set_skip_taskbar_hint(true);
        gtk_window.set_keep_above(true);
        gtk_window.set_accept_focus(false);
        gtk_window.set_focus_on_map(false);
    });
    Ok(window)
}

/// The pasting strip has the keyboard (its losing it hides the strip).
static TOOK_KEYBOARD: std::sync::atomic::AtomicBool = std::sync::atomic::AtomicBool::new(false);

/// Work on the strip's GTK window, on the main thread (waits for it, at most a second).
fn with_gtk(window: &WebviewWindow, work: impl FnOnce(&gtk::ApplicationWindow) + Send + 'static) {
    let Some(app) = app() else { return };
    let window = window.clone();
    let (sender, receiver) = std::sync::mpsc::sync_channel(1);
    let scheduled = app.run_on_main_thread(move || {
        if let Ok(gtk_window) = window.gtk_window() {
            work(&gtk_window);
        }
        let _ = sender.send(());
    });
    if scheduled.is_ok() {
        let _ = receiver.recv_timeout(Duration::from_secs(1));
    }
}

/// The mouse pointer on the screen (pixels), where a formula opened without a caret stands.
pub(super) fn pointer() -> Option<(f64, f64)> {
    let app = app()?;
    let (sender, receiver) = std::sync::mpsc::sync_channel(1);
    app.run_on_main_thread(move || {
        let position = gdk::Display::default()
            .and_then(|display| display.default_seat())
            .and_then(|seat| seat.pointer())
            .map(|pointer| {
                let (_, x, y) = pointer.position();
                (x as f64, y as f64)
            });
        let _ = sender.send(position);
    })
    .ok()?;
    receiver.recv_timeout(Duration::from_millis(500)).ok()?
}

pub(super) fn window_number(_window: &WebviewWindow) -> Option<u32> {
    Some(0)
}

/// The mouse goes through the window (the input method's formula) or not (the pasting strip).
fn set_click_through(gtk_window: &gtk::ApplicationWindow, through: bool) {
    if through {
        let region = gtk::cairo::Region::create();
        gtk_window.input_shape_combine_region(Some(&region));
    } else {
        gtk_window.input_shape_combine_region(None);
    }
}

pub(super) fn park(window: &WebviewWindow) {
    TOOK_KEYBOARD.store(false, std::sync::atomic::Ordering::SeqCst);
    let _ = window.hide();
}

pub(super) fn park_for_reload(window: &WebviewWindow) {
    park(window);
}

/// Hides the strip and returns once the window manager has had a moment to give the keyboard back.
pub(super) fn park_and_wait(window: &WebviewWindow) -> Result<(), String> {
    park(window);
    std::thread::sleep(Duration::from_millis(80));
    Ok(())
}

pub(super) fn present(window: &WebviewWindow, anchor: ScreenRect, layout: PanelLayout) {
    place(window, anchor, layout, Placement::OnLine);
    with_gtk(window, |gtk_window| {
        set_click_through(gtk_window, false);
        gtk_window.set_accept_focus(true);
    });
    let _ = window.show();
    let _ = window.set_focus();
    TOOK_KEYBOARD.store(true, std::sync::atomic::Ordering::SeqCst);
}

/// The input method's formula: shown without taking the keyboard or the mouse.
pub(super) fn present_passive(window: &WebviewWindow, anchor: ScreenRect, layout: PanelLayout, placement: Placement) {
    TOOK_KEYBOARD.store(false, std::sync::atomic::Ordering::SeqCst);
    place(window, anchor, layout, placement);
    with_gtk(window, |gtk_window| {
        gtk_window.set_accept_focus(false);
        set_click_through(gtk_window, true);
    });
    let _ = window.show();
    // shown windows can be moved by the window manager: put it back where it goes
    place(window, anchor, layout, placement);
}

pub(super) fn reframe(window: &WebviewWindow, anchor: ScreenRect, layout: PanelLayout, placement: Placement) {
    place(window, anchor, layout, placement);
}

/// The input method's places are the screen's own (y down): nothing to flip.
pub(super) fn main_display_height() -> f64 {
    0.0
}

fn place(window: &WebviewWindow, anchor: ScreenRect, layout: PanelLayout, placement: Placement) {
    let scale = window.scale_factor().unwrap_or(1.0);
    let (x, y, width, height) = frame_for(anchor, layout, placement, scale);
    let _ = window.set_size(tauri::PhysicalSize::new(width.round().max(1.0) as u32, height.round().max(1.0) as u32));
    let _ = window.set_position(tauri::PhysicalPosition::new(x.round() as i32, y.round() as i32));
}

/// The window's frame (screen pixels) putting the strip's anchor on the caret's vertical centre
/// (OnLine), its top edge under the caret line (BelowLine) or the formula's origin on the line's
/// baseline (InLine), kept on the caret's monitor except in the line.
fn frame_for(anchor: ScreenRect, layout: PanelLayout, placement: Placement, scale: f64) -> (f64, f64, f64, f64) {
    let (width, height) = (layout.width * scale, layout.height * scale);
    let mut x = anchor.x - layout.anchor_x * scale;
    let mut y = match placement {
        Placement::OnLine => anchor.y + anchor.height / 2.0 - layout.anchor_y * scale,
        Placement::BelowLine => anchor.y + anchor.height + BELOW_LINE_GAP - layout.anchor_y * scale,
        Placement::InLine => anchor.y - layout.anchor_y * scale,
    };
    if placement != Placement::InLine {
        if let Some((left, top, right, bottom)) = monitor_at(anchor.x, anchor.y) {
            x = x.min(right - width).max(left);
            if placement == Placement::BelowLine && y + height > bottom {
                // no room under the line: the card goes over it
                y = anchor.y - BELOW_LINE_GAP - height;
            }
            y = y.min(bottom - height).max(top);
        }
    }
    (x, y, width, height)
}

/// The work area of the monitor at a point (pixels): left, top, right, bottom.
fn monitor_at(x: f64, y: f64) -> Option<(f64, f64, f64, f64)> {
    let app = app()?;
    let (sender, receiver) = std::sync::mpsc::sync_channel(1);
    app.run_on_main_thread(move || {
        let area = gdk::Display::default().and_then(|display| {
            let monitor = display.monitor_at_point(x as i32, y as i32)?;
            let area = monitor.workarea();
            let scale = monitor.scale_factor() as f64;
            Some((
                area.x() as f64 * scale,
                area.y() as f64 * scale,
                (area.x() + area.width()) as f64 * scale,
                (area.y() + area.height()) as f64 * scale,
            ))
        });
        let _ = sender.send(area);
    })
    .ok()?;
    receiver.recv_timeout(Duration::from_millis(500)).ok()?
}
