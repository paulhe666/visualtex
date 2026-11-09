//! What a formula opened without the input method stands on, on Linux: X11 tells no caret to other
//! applications, so the mouse pointer (the input method reports the caret itself).

use super::{AnchorSource, ScreenRect, Target};

pub(super) fn is_trusted() -> bool {
    true
}

pub(super) fn request_trust() {}

pub(super) fn capture_target(_trusted: bool) -> Target {
    let (x, y) = super::panel::pointer().unwrap_or((240.0, 240.0));
    Target {
        app_name: String::new(),
        anchor: ScreenRect { x, y: y - 9.0, width: 1.0, height: 18.0 },
        source: AnchorSource::Pointer,
    }
}

pub(super) fn frontmost_bundle() -> Option<String> {
    None
}

pub(super) fn frontmost_name() -> Option<String> {
    None
}
