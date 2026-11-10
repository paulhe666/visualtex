//! The pasting strip's delivery on Linux: the formula on the clipboard, then Ctrl+V into the field
//! (XTest, X11); and the clipboard's image or text for a formula (Ctrl+V in it).

use gtk::prelude::*;
use std::time::Duration;
use tauri::AppHandle;

fn on_main<T: Send + 'static>(app: &AppHandle, work: impl FnOnce() -> T + Send + 'static) -> Result<T, String> {
    let (sender, receiver) = std::sync::mpsc::sync_channel(1);
    app.run_on_main_thread(move || {
        let _ = sender.send(work());
    })
    .map_err(|error| format!("Unable to reach the main thread: {error}"))?;
    receiver
        .recv_timeout(Duration::from_secs(5))
        .map_err(|_| "The clipboard did not answer".to_string())
}

pub(super) fn copy(app: &AppHandle, text: String) -> Result<(), String> {
    on_main(app, move || {
        let clipboard = gtk::Clipboard::get(&gdk::SELECTION_CLIPBOARD);
        clipboard.set_text(&text);
        clipboard.store();
    })
}

/// Ctrl+V into the window with the keyboard (the field the formula was opened in), pressed
/// through XTest (X11; a Wayland application does not take it: the formula stays on the clipboard).
fn press_paste() -> Result<(), String> {
    use x11rb::connection::Connection;
    use x11rb::protocol::xproto::{ConnectionExt as _, KEY_PRESS_EVENT, KEY_RELEASE_EVENT};
    use x11rb::protocol::xtest::ConnectionExt as _;
    const COPIED: &str = "The formula is on the clipboard: paste it with Ctrl+V";
    std::thread::sleep(Duration::from_millis(60));
    let (connection, screen) = x11rb::connect(None).map_err(|_| COPIED.to_string())?;
    let root = connection.setup().roots.get(screen).ok_or(COPIED)?.root;
    let (min, max) = (connection.setup().min_keycode, connection.setup().max_keycode);
    let mapping = connection
        .get_keyboard_mapping(min, max - min + 1)
        .map_err(|_| COPIED.to_string())?
        .reply()
        .map_err(|_| COPIED.to_string())?;
    let per = mapping.keysyms_per_keycode.max(1) as usize;
    let keycode = |keysym: u32| {
        mapping
            .keysyms
            .chunks(per)
            .position(|keysyms| keysyms.first() == Some(&keysym))
            .map(|index| min + index as u8)
    };
    // Control_L and v
    let (Some(control), Some(v)) = (keycode(0xffe3), keycode(0x76)) else { return Err(COPIED.to_string()) };
    let fake = |kind: u8, key: u8| connection.xtest_fake_input(kind, key, 0, root, 0, 0, 0).map(|_| ());
    fake(KEY_PRESS_EVENT, control)
        .and_then(|_| fake(KEY_PRESS_EVENT, v))
        .and_then(|_| fake(KEY_RELEASE_EVENT, v))
        .and_then(|_| fake(KEY_RELEASE_EVENT, control))
        .map_err(|_| COPIED.to_string())?;
    connection.flush().map_err(|_| COPIED.to_string())
}

pub(super) fn paste(app: &AppHandle, text: String) -> Result<(), String> {
    copy(app, text)?;
    press_paste()
}

pub(super) fn paste_png(app: &AppHandle, png: Vec<u8>) -> Result<(), String> {
    on_main(app, move || -> Result<(), String> {
        let loader = gdk_pixbuf::PixbufLoader::new();
        loader.write(&png).map_err(|error| error.to_string())?;
        loader.close().map_err(|error| error.to_string())?;
        let pixbuf = loader.pixbuf().ok_or("The formula image could not be read")?;
        let clipboard = gtk::Clipboard::get(&gdk::SELECTION_CLIPBOARD);
        clipboard.set_image(&pixbuf);
        clipboard.store();
        Ok(())
    })??;
    press_paste()
}

pub(super) enum ClipboardContent {
    Image { bytes: Vec<u8>, extension: String },
    Text(String),
    Nothing,
}

pub(super) fn read_clipboard(app: &AppHandle) -> Result<ClipboardContent, String> {
    on_main(app, || {
        let clipboard = gtk::Clipboard::get(&gdk::SELECTION_CLIPBOARD);
        if let Some(image) = clipboard.wait_for_image() {
            if let Ok(bytes) = image.save_to_bufferv("png", &[]) {
                return ClipboardContent::Image { bytes, extension: "png".to_string() };
            }
        }
        match clipboard.wait_for_text() {
            Some(text) if !text.is_empty() => ClipboardContent::Text(text.to_string()),
            _ => ClipboardContent::Nothing,
        }
    })
}
