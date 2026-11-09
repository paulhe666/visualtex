//! The probe key. Chromium (and Electron) on X11 tell the input method where their cursor is only
//! while handling a key event; to read where a field put the formula's pieces there, a key no
//! keyboard has (a key code without symbols) is pressed through XTest after each move of the
//! preedit's cursor, and the engine swallows it when it comes through.

use std::sync::atomic::{AtomicU32, Ordering};
use std::sync::Mutex;
use x11rb::connection::Connection;
use x11rb::protocol::xproto::{ConnectionExt as _, KEY_PRESS_EVENT, KEY_RELEASE_EVENT};
use x11rb::protocol::xtest::ConnectionExt as _;
use x11rb::rust_connection::RustConnection;

/// The probe key's code as IBus reports it (the X key code less 8), once known.
static PROBE_KEYCODE: AtomicU32 = AtomicU32::new(0);

struct Display {
    connection: RustConnection,
    root: u32,
    keycode: u8,
}

static DISPLAY: Mutex<Option<Option<Display>>> = Mutex::new(None);

fn open() -> Option<Display> {
    std::env::var_os("DISPLAY")?;
    let (connection, screen) = x11rb::connect(None).ok()?;
    let root = connection.setup().roots.get(screen)?.root;
    let (min, max) = (connection.setup().min_keycode, connection.setup().max_keycode);
    let mapping = connection.get_keyboard_mapping(min, max - min + 1).ok()?.reply().ok()?;
    let per = mapping.keysyms_per_keycode.max(1) as usize;
    // the highest key code without symbols: no key on the keyboard, nothing in any application
    let keycode = mapping
        .keysyms
        .chunks(per)
        .rposition(|keysyms| keysyms.iter().all(|keysym| *keysym == 0))
        .map(|index| min + index as u8)?;
    PROBE_KEYCODE.store(u32::from(keycode).saturating_sub(8), Ordering::SeqCst);
    crate::log::write(&format!("probe key: keycode {keycode}"));
    Some(Display { connection, root, keycode })
}

/// Presses and releases the probe key (into the window with the keyboard); false without X11.
pub fn press_probe_key() -> bool {
    let Ok(mut display) = DISPLAY.lock() else { return false };
    if display.is_none() {
        *display = Some(open());
    }
    let Some(Some(display)) = display.as_ref() else { return false };
    display
        .connection
        .xtest_fake_input(KEY_PRESS_EVENT, display.keycode, 0, display.root, 0, 0, 0)
        .and_then(|_| display.connection.xtest_fake_input(KEY_RELEASE_EVENT, display.keycode, 0, display.root, 0, 0, 0))
        .is_ok()
        && display.connection.flush().is_ok()
}

/// The probe key reached the engine (by its key code): swallowed.
pub fn is_probe_key(keycode: u32) -> bool {
    let probe = PROBE_KEYCODE.load(Ordering::SeqCst);
    probe != 0 && keycode == probe
}
