//! Where the user is typing: the frontmost application's focused text
//! element and its caret rectangle, read through Accessibility.
//!
//! Every call into another application waits at most AX_TIMEOUT_SECONDS (the
//! system default is about six seconds). Without the Accessibility permission,
//! or when an application reports no caret, the strip falls back to the
//! focused field and then to the mouse pointer.

use super::{AnchorSource, ScreenRect, Target};
use objc2_app_kit::NSWorkspace;
use std::ffi::c_void;
use std::ptr;
use std::sync::atomic::{AtomicBool, Ordering};

type CFTypeRef = *const c_void;
type AXError = i32;

const AX_SUCCESS: AXError = 0;
const AX_TIMEOUT_SECONDS: f32 = 0.25;
const AX_VALUE_CG_POINT: u32 = 1;
const AX_VALUE_CG_SIZE: u32 = 2;
const AX_VALUE_CG_RECT: u32 = 3;
const AX_VALUE_CF_RANGE: u32 = 4;
const CF_STRING_ENCODING_UTF8: u32 = 0x0800_0100;
/// The line height assumed when nothing reports one.
const FALLBACK_LINE_HEIGHT: f64 = 18.0;

#[repr(C)]
#[derive(Clone, Copy, Default)]
struct CGPoint {
    x: f64,
    y: f64,
}

#[repr(C)]
#[derive(Clone, Copy, Default)]
struct CGSize {
    width: f64,
    height: f64,
}

#[repr(C)]
#[derive(Clone, Copy, Default)]
struct CGRect {
    origin: CGPoint,
    size: CGSize,
}

#[repr(C)]
#[derive(Clone, Copy, Default)]
struct CFRange {
    location: isize,
    length: isize,
}

#[repr(C)]
struct CFDictionaryCallBacks {
    _opaque: [u8; 0],
}

#[link(name = "ApplicationServices", kind = "framework")]
unsafe extern "C" {
    fn AXIsProcessTrusted() -> u8;
    fn AXIsProcessTrustedWithOptions(options: CFTypeRef) -> u8;
    static kAXTrustedCheckOptionPrompt: CFTypeRef;
    fn AXUIElementCreateApplication(pid: i32) -> CFTypeRef;
    fn AXUIElementCopyAttributeValue(element: CFTypeRef, attribute: CFTypeRef, value: *mut CFTypeRef) -> AXError;
    fn AXUIElementCopyParameterizedAttributeValue(
        element: CFTypeRef,
        attribute: CFTypeRef,
        parameter: CFTypeRef,
        value: *mut CFTypeRef,
    ) -> AXError;
    fn AXUIElementSetAttributeValue(element: CFTypeRef, attribute: CFTypeRef, value: CFTypeRef) -> AXError;
    fn AXUIElementSetMessagingTimeout(element: CFTypeRef, timeout: f32) -> AXError;
    fn AXUIElementGetTypeID() -> usize;
    fn AXValueCreate(value_type: u32, value: *const c_void) -> CFTypeRef;
    fn AXValueGetType(value: CFTypeRef) -> u32;
    fn AXValueGetValue(value: CFTypeRef, value_type: u32, out: *mut c_void) -> u8;
}

#[link(name = "CoreFoundation", kind = "framework")]
unsafe extern "C" {
    fn CFRelease(value: CFTypeRef);
    fn CFGetTypeID(value: CFTypeRef) -> usize;
    fn CFStringCreateWithBytes(
        allocator: CFTypeRef,
        bytes: *const u8,
        length: isize,
        encoding: u32,
        external_representation: u8,
    ) -> CFTypeRef;
    fn CFDictionaryCreate(
        allocator: CFTypeRef,
        keys: *const CFTypeRef,
        values: *const CFTypeRef,
        count: isize,
        key_callbacks: *const CFDictionaryCallBacks,
        value_callbacks: *const CFDictionaryCallBacks,
    ) -> CFTypeRef;
    static kCFTypeDictionaryKeyCallBacks: CFDictionaryCallBacks;
    static kCFTypeDictionaryValueCallBacks: CFDictionaryCallBacks;
    static kCFBooleanTrue: CFTypeRef;
}

#[link(name = "CoreGraphics", kind = "framework")]
unsafe extern "C" {
    fn CGEventCreate(source: CFTypeRef) -> CFTypeRef;
    fn CGEventGetLocation(event: CFTypeRef) -> CGPoint;
}

/// A Core Foundation object this module owns (released on drop).
struct Owned(CFTypeRef);

impl Owned {
    fn new(value: CFTypeRef) -> Option<Self> {
        (!value.is_null()).then_some(Self(value))
    }
}

impl Drop for Owned {
    fn drop(&mut self) {
        unsafe { CFRelease(self.0) }
    }
}

fn cf_string(value: &str) -> Option<Owned> {
    Owned::new(unsafe {
        CFStringCreateWithBytes(
            ptr::null(),
            value.as_ptr(),
            value.len() as isize,
            CF_STRING_ENCODING_UTF8,
            0,
        )
    })
}

pub(super) fn is_trusted() -> bool {
    unsafe { AXIsProcessTrusted() != 0 }
}

/// Shows macOS's "VisualTeX would like to control this computer" prompt, at
/// most once per launch.
pub(super) fn request_trust() {
    static ASKED: AtomicBool = AtomicBool::new(false);
    if ASKED.swap(true, Ordering::SeqCst) {
        return;
    }
    unsafe {
        let keys = [kAXTrustedCheckOptionPrompt];
        let values = [kCFBooleanTrue];
        let Some(options) = Owned::new(CFDictionaryCreate(
            ptr::null(),
            keys.as_ptr(),
            values.as_ptr(),
            1,
            &kCFTypeDictionaryKeyCallBacks,
            &kCFTypeDictionaryValueCallBacks,
        )) else {
            return;
        };
        AXIsProcessTrustedWithOptions(options.0);
    }
}

pub(super) fn capture_target(trusted: bool) -> Target {
    let (pid, app_name) = frontmost_application();
    if trusted && pid > 0 {
        if let Some((anchor, source)) = focused_text_anchor(pid) {
            return Target { app_name, anchor, source };
        }
    }
    Target {
        app_name,
        anchor: pointer_anchor(),
        source: AnchorSource::Pointer,
    }
}

fn frontmost_application() -> (i32, String) {
    let Some(application) = NSWorkspace::sharedWorkspace().frontmostApplication() else {
        return (0, String::new());
    };
    let name = application
        .localizedName()
        .map(|name| name.to_string())
        .unwrap_or_default();
    (application.processIdentifier(), name)
}

fn copy_attribute(element: &Owned, attribute: &str) -> Option<Owned> {
    let name = cf_string(attribute)?;
    let mut value: CFTypeRef = ptr::null();
    let error = unsafe { AXUIElementCopyAttributeValue(element.0, name.0, &mut value) };
    if error != AX_SUCCESS {
        return None;
    }
    Owned::new(value)
}

fn copy_element(element: &Owned, attribute: &str) -> Option<Owned> {
    copy_attribute(element, attribute)
        .filter(|value| unsafe { CFGetTypeID(value.0) == AXUIElementGetTypeID() })
}

fn ax_value<T: Default>(value: &Owned, value_type: u32) -> Option<T> {
    unsafe {
        if AXValueGetType(value.0) != value_type {
            return None;
        }
        let mut out = T::default();
        (AXValueGetValue(value.0, value_type, (&mut out as *mut T).cast()) != 0).then_some(out)
    }
}

fn set_true(element: &Owned, attribute: &str) {
    if let Some(name) = cf_string(attribute) {
        unsafe {
            AXUIElementSetAttributeValue(element.0, name.0, kCFBooleanTrue);
        }
    }
}

fn bounds_for_range(element: &Owned, range: CFRange) -> Option<ScreenRect> {
    let parameter = Owned::new(unsafe { AXValueCreate(AX_VALUE_CF_RANGE, (&range as *const CFRange).cast()) })?;
    let name = cf_string("AXBoundsForRange")?;
    let mut value: CFTypeRef = ptr::null();
    let error = unsafe { AXUIElementCopyParameterizedAttributeValue(element.0, name.0, parameter.0, &mut value) };
    if error != AX_SUCCESS {
        return None;
    }
    let rect: CGRect = ax_value(&Owned::new(value)?, AX_VALUE_CG_RECT)?;
    let rect = ScreenRect {
        x: rect.origin.x,
        y: rect.origin.y,
        width: rect.size.width,
        height: rect.size.height,
    };
    is_plausible_line(&rect).then_some(rect)
}

/// Some applications answer an empty range with a zero rectangle, or with the
/// whole document.
fn is_plausible_line(rect: &ScreenRect) -> bool {
    [rect.x, rect.y, rect.width, rect.height].iter().all(|value| value.is_finite())
        && rect.width >= 0.0
        && (4.0..=320.0).contains(&rect.height)
        && !(rect.x == 0.0 && rect.y == 0.0)
}

/// The text of a range, to tell a line break before the caret.
fn string_for_range(element: &Owned, range: CFRange) -> Option<String> {
    let parameter = Owned::new(unsafe { AXValueCreate(AX_VALUE_CF_RANGE, (&range as *const CFRange).cast()) })?;
    let name = cf_string("AXStringForRange")?;
    let mut value: CFTypeRef = ptr::null();
    let error = unsafe { AXUIElementCopyParameterizedAttributeValue(element.0, name.0, parameter.0, &mut value) };
    if error != AX_SUCCESS {
        return None;
    }
    cf_string_value(&Owned::new(value)?)
}

fn cf_string_value(value: &Owned) -> Option<String> {
    #[link(name = "CoreFoundation", kind = "framework")]
    unsafe extern "C" {
        fn CFStringGetTypeID() -> usize;
        fn CFStringGetCString(value: CFTypeRef, buffer: *mut i8, size: isize, encoding: u32) -> u8;
    }
    unsafe {
        if CFGetTypeID(value.0) != CFStringGetTypeID() {
            return None;
        }
        let mut buffer = [0i8; 256];
        if CFStringGetCString(value.0, buffer.as_mut_ptr(), buffer.len() as isize, CF_STRING_ENCODING_UTF8) == 0 {
            return None;
        }
        Some(std::ffi::CStr::from_ptr(buffer.as_ptr()).to_string_lossy().into_owned())
    }
}

fn is_line_break(text: &str) -> bool {
    matches!(text, "\n" | "\r" | "\r\n" | "\u{2028}" | "\u{2029}")
}

/// The caret as a zero-width rectangle on its line. Character rectangles are
/// what applications report most reliably: an empty range is answered one
/// line too high by TextEdit (its first line read 90 for a text area starting
/// at 103) and with nothing by others. So the caret stands at the right edge
/// of the character before it, or at a line start on the left edge of the
/// character after it; only without either the empty range counts. A
/// selection anchors at its first character: the formula replaces it from
/// there.
fn caret_rect(element: &Owned, range: CFRange) -> Option<ScreenRect> {
    let at = |rect: ScreenRect, x: f64| ScreenRect { x, width: 0.0, ..rect };
    if range.length > 0 {
        let first = CFRange { location: range.location, length: 1 };
        return bounds_for_range(element, first)
            .or_else(|| bounds_for_range(element, range))
            .map(|rect| at(rect, rect.x));
    }
    let after = CFRange { location: range.location, length: 1 };
    if range.location > 0 {
        let before = CFRange { location: range.location - 1, length: 1 };
        let line_start = string_for_range(element, before).is_some_and(|text| is_line_break(&text));
        if !line_start {
            if let Some(rect) = bounds_for_range(element, before) {
                return Some(at(rect, rect.x + rect.width));
            }
        }
    }
    let at_line_break = string_for_range(element, after).is_some_and(|text| is_line_break(&text));
    if !at_line_break {
        if let Some(rect) = bounds_for_range(element, after) {
            return Some(at(rect, rect.x));
        }
    }
    bounds_for_range(element, range).map(|rect| at(rect, rect.x))
}

fn element_frame(element: &Owned) -> Option<ScreenRect> {
    let position: CGPoint = ax_value(&copy_attribute(element, "AXPosition")?, AX_VALUE_CG_POINT)?;
    let size: CGSize = ax_value(&copy_attribute(element, "AXSize")?, AX_VALUE_CG_SIZE)?;
    let frame = ScreenRect {
        x: position.x,
        y: position.y,
        width: size.width,
        height: size.height,
    };
    ([frame.x, frame.y, frame.width, frame.height].iter().all(|value| value.is_finite())
        && frame.width > 0.0
        && frame.height > 0.0)
        .then_some(frame)
}

/// The application's focused element. Chromium browsers (Chrome, Edge) and
/// Electron applications build their accessibility tree only once an
/// assistive application asks for it, and it arrives a moment later; Chrome
/// answers AXEnhancedUserInterface, Electron AXManualAccessibility. Both stay
/// on, so only the first activation in such an application waits.
fn focused_element(application: &Owned) -> Option<Owned> {
    if let Some(element) = copy_element(application, "AXFocusedUIElement") {
        return Some(element);
    }
    for attribute in ["AXManualAccessibility", "AXEnhancedUserInterface"] {
        set_true(application, attribute);
        for _ in 0..4 {
            std::thread::sleep(std::time::Duration::from_millis(30));
            if let Some(element) = copy_element(application, "AXFocusedUIElement") {
                super::log(&format!("ax focused element after {attribute}"));
                return Some(element);
            }
        }
    }
    super::log("ax no focused element");
    None
}

fn focused_text_anchor(pid: i32) -> Option<(ScreenRect, AnchorSource)> {
    let application = Owned::new(unsafe { AXUIElementCreateApplication(pid) })?;
    unsafe {
        AXUIElementSetMessagingTimeout(application.0, AX_TIMEOUT_SECONDS);
    }
    let focused = focused_element(&application)?;
    unsafe {
        AXUIElementSetMessagingTimeout(focused.0, AX_TIMEOUT_SECONDS);
    }
    let range = copy_attribute(&focused, "AXSelectedTextRange")
        .and_then(|value| ax_value::<CFRange>(&value, AX_VALUE_CF_RANGE));
    log_geometry(&focused, range);
    if let Some(range) = range {
        if let Some(caret) = caret_rect(&focused, range) {
            return Some((caret, AnchorSource::Caret));
        }
    }
    // No caret: stand at the start of the field's first line, where an input
    // method's composition window goes when the field reports nothing.
    let frame = element_frame(&focused)?;
    let line = frame.height.clamp(14.0, 24.0);
    Some((
        ScreenRect {
            x: frame.x + 4.0,
            y: frame.y + ((frame.height - line) / 2.0).clamp(0.0, 4.0),
            width: 0.0,
            height: line,
        },
        AnchorSource::Field,
    ))
}

fn string_attribute(element: &Owned, attribute: &str) -> String {
    copy_attribute(element, attribute)
        .and_then(|value| cf_string_value(&value))
        .unwrap_or_default()
}

/// What the focused element reports, for anywhere.log: applications differ
/// in how they answer an empty range.
fn log_geometry(element: &Owned, range: Option<CFRange>) {
    let rect = |range: CFRange| {
        bounds_for_range(element, range)
            .map(|r| format!("{:.1},{:.1} {:.1}x{:.1}", r.x, r.y, r.width, r.height))
            .unwrap_or_else(|| "-".to_string())
    };
    let frame = element_frame(element)
        .map(|r| format!("{:.1},{:.1} {:.1}x{:.1}", r.x, r.y, r.width, r.height))
        .unwrap_or_else(|| "-".to_string());
    let detail = match range {
        Some(range) => format!(
            "range={},{} empty=[{}] before=[{}] after=[{}]",
            range.location,
            range.length,
            rect(range),
            if range.location > 0 { rect(CFRange { location: range.location - 1, length: 1 }) } else { "-".to_string() },
            rect(CFRange { location: range.location, length: 1 }),
        ),
        None => "range=-".to_string(),
    };
    super::log(&format!(
        "ax role={} subrole={} frame=[{frame}] {detail}",
        string_attribute(element, "AXRole"),
        string_attribute(element, "AXSubrole"),
    ));
}

fn pointer_anchor() -> ScreenRect {
    let location = Owned::new(unsafe { CGEventCreate(ptr::null()) })
        .map(|event| unsafe { CGEventGetLocation(event.0) })
        .unwrap_or_default();
    ScreenRect {
        x: location.x,
        y: location.y - FALLBACK_LINE_HEIGHT / 2.0,
        width: 0.0,
        height: FALLBACK_LINE_HEIGHT,
    }
}
