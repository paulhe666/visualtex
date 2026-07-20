use serde::{Deserialize, Serialize};
use std::collections::VecDeque;
#[cfg(target_os = "macos")]
use std::fs;
use std::io::Read;
#[cfg(target_os = "macos")]
use std::path::PathBuf;
use std::process::{Command, Stdio};
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

const POWERPOINT_BUNDLE_ID: &str = "com.microsoft.Powerpoint";
const WORD_BUNDLE_ID: &str = "com.microsoft.Word";
const SHAPE_PREFIX: &str = "VisualTeX_";
const WORD_METADATA_PREFIX: &str = "visualtex:v1:deflate:";
const WORD_SELECTION_FIELD_SEPARATOR: &str = "<VISUALTEX_WORD_FIELD>";
const POWERPOINT_SNAPSHOT_FIELD_SEPARATOR: &str = "<VISUALTEX_PPT_FIELD>";
const POWERPOINT_SNAPSHOT_RECORD_SEPARATOR: &str = "<VISUALTEX_PPT_RECORD>";
const MAX_EVENTS: usize = 64;
const APPLESCRIPT_TIMEOUT: Duration = Duration::from_secs(20);
const APPLESCRIPT_QUERY_TIMEOUT: Duration = Duration::from_secs(3);
const APPLESCRIPT_LOCK_TIMEOUT: Duration = Duration::from_secs(3);
static POWERPOINT_APPLESCRIPT_LOCK: Mutex<()> = Mutex::new(());

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct PowerPointNativeSelection {
    pub shape_name: String,
    pub slide_index: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub slide_id: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub presentation_identity: Option<String>,
    pub left: f64,
    pub top: f64,
    pub width: f64,
    pub height: f64,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
struct WordNativeFormulaSelection {
    marker: String,
    width: f64,
    height: f64,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct PowerPointNativeSlideSnapshot {
    pub presentation_identity: String,
    pub slide_index: u32,
    pub slide_id: u32,
    pub shape_count: u32,
    pub shape_names: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PowerPointInteractionEvent {
    pub cursor: u64,
    pub host: &'static str,
    pub kind: &'static str,
    pub formula_id: String,
    pub shape_name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub slide_index: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub slide_id: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub presentation_identity: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    pub created_at: u64,
}

#[derive(Debug, Default)]
struct InteractionState {
    next_cursor: u64,
    delivered_powerpoint_cursor: u64,
    delivered_word_cursor: u64,
    events: VecDeque<PowerPointInteractionEvent>,
}

#[derive(Debug, Clone, Default)]
pub struct PowerPointInteractionBus {
    inner: Arc<Mutex<InteractionState>>,
}

impl PowerPointInteractionBus {
    #[cfg(test)]
    pub fn push_edit_selected(&self, host: &'static str, shape_name: String, formula_id: String) {
        self.push_edit_target("edit-selected", host, shape_name, formula_id, None);
    }

    pub fn push_powerpoint_edit_selected(
        &self,
        selection: PowerPointNativeSelection,
        formula_id: String,
    ) {
        let shape_name = selection.shape_name.clone();
        self.push_edit_target(
            "edit-selected",
            "powerpoint",
            shape_name,
            formula_id,
            Some(selection),
        );
    }

    pub fn push_powerpoint_edit_requested(&self, selection: PowerPointNativeSelection) {
        let shape_name = selection.shape_name.clone();
        self.push_edit_target(
            "edit-requested",
            "powerpoint",
            shape_name,
            String::new(),
            Some(selection),
        );
    }

    fn push_word_edit_selected(&self, selection: WordNativeFormulaSelection) {
        let marker = selection.marker.clone();
        self.push_edit_target(
            "edit-selected",
            "word",
            marker.clone(),
            String::new(),
            Some(PowerPointNativeSelection {
                shape_name: marker,
                slide_index: 0,
                slide_id: None,
                presentation_identity: None,
                left: 0.0,
                top: 0.0,
                width: selection.width,
                height: selection.height,
            }),
        );
    }

    fn push_edit_target(
        &self,
        kind: &'static str,
        host: &'static str,
        shape_name: String,
        formula_id: String,
        selection: Option<PowerPointNativeSelection>,
    ) {
        let created_at = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|value| value.as_millis() as u64)
            .unwrap_or_default();
        if let Ok(mut state) = self.inner.lock() {
            state.next_cursor = state.next_cursor.saturating_add(1);
            let cursor = state.next_cursor;
            state.events.push_back(PowerPointInteractionEvent {
                cursor,
                host,
                kind,
                formula_id,
                shape_name,
                slide_index: selection.as_ref().map(|value| value.slide_index),
                slide_id: selection.as_ref().and_then(|value| value.slide_id),
                presentation_identity: selection
                    .as_ref()
                    .and_then(|value| value.presentation_identity.clone()),
                left: selection.as_ref().map(|value| value.left),
                top: selection.as_ref().map(|value| value.top),
                width: selection.as_ref().map(|value| value.width),
                height: selection.as_ref().map(|value| value.height),
                created_at,
            });
            while state.events.len() > MAX_EVENTS {
                state.events.pop_front();
            }
        }
    }

    pub fn take_after(&self, host: &'static str, cursor: u64) -> Vec<PowerPointInteractionEvent> {
        self.inner
            .lock()
            .map(|mut state| {
                let delivered_cursor = match host {
                    "powerpoint" => state.delivered_powerpoint_cursor,
                    "word" => state.delivered_word_cursor,
                    _ => return Vec::new(),
                };
                let threshold = cursor.max(delivered_cursor);
                let events = state
                    .events
                    .iter()
                    .filter(|event| event.host == host && event.cursor > threshold)
                    .cloned()
                    .collect::<Vec<_>>();
                if let Some(last) = events.last() {
                    match host {
                        "powerpoint" => state.delivered_powerpoint_cursor = last.cursor,
                        "word" => state.delivered_word_cursor = last.cursor,
                        _ => {}
                    }
                }
                events
            })
            .unwrap_or_default()
    }
}

fn run_applescript(script: &str) -> Result<String, String> {
    run_applescript_with_timeout(script, APPLESCRIPT_TIMEOUT)
}

fn run_applescript_with_timeout(script: &str, timeout: Duration) -> Result<String, String> {
    #[cfg(not(target_os = "macos"))]
    {
        let _ = script;
        return Err("PowerPoint native integration is only available on macOS".to_string());
    }

    #[cfg(target_os = "macos")]
    {
        let lock_deadline = Instant::now() + APPLESCRIPT_LOCK_TIMEOUT;
        let _operation_guard = loop {
            match POWERPOINT_APPLESCRIPT_LOCK.try_lock() {
                Ok(guard) => break guard,
                Err(std::sync::TryLockError::WouldBlock) if Instant::now() < lock_deadline => {
                    thread::sleep(Duration::from_millis(25));
                }
                Err(std::sync::TryLockError::WouldBlock) => {
                    return Err(
                        "PowerPoint is still finishing another VisualTeX operation. Try again."
                            .to_string(),
                    );
                }
                Err(std::sync::TryLockError::Poisoned(_)) => {
                    return Err("PowerPoint automation lock is unavailable".to_string());
                }
            }
        };
        let mut child = Command::new("/usr/bin/osascript")
            .arg("-e")
            .arg(script)
            .stdout(Stdio::piped())
            .stderr(Stdio::piped())
            .spawn()
            .map_err(|error| format!("Unable to launch AppleScript: {error}"))?;
        let deadline = Instant::now() + timeout;
        let status = loop {
            match child.try_wait() {
                Ok(Some(status)) => break status,
                Ok(None) if Instant::now() < deadline => {
                    thread::sleep(Duration::from_millis(25));
                }
                Ok(None) => {
                    let _ = child.kill();
                    let _ = child.wait();
                    return Err(
                        "PowerPoint operation timed out. Close any open PowerPoint dialog and try again."
                            .to_string(),
                    );
                }
                Err(error) => {
                    let _ = child.kill();
                    let _ = child.wait();
                    return Err(format!("Unable to monitor PowerPoint AppleScript: {error}"));
                }
            }
        };
        let mut stdout = Vec::new();
        let mut stderr = Vec::new();
        if let Some(mut pipe) = child.stdout.take() {
            let _ = pipe.read_to_end(&mut stdout);
        }
        if let Some(mut pipe) = child.stderr.take() {
            let _ = pipe.read_to_end(&mut stderr);
        }
        if !status.success() {
            let detail = String::from_utf8_lossy(&stderr).trim().to_string();
            return Err(if detail.is_empty() {
                "PowerPoint AppleScript operation failed".to_string()
            } else {
                detail
            });
        }
        Ok(String::from_utf8_lossy(&stdout).trim().to_string())
    }
}

fn parse_selection(output: &str) -> Result<PowerPointNativeSelection, String> {
    let fields: Vec<&str> = output.split('\u{1f}').collect();
    if fields.len() != 6 && fields.len() != 8 {
        return Err(format!(
            "PowerPoint returned an invalid selection payload: {output}"
        ));
    }
    let parse_number = |value: &str, label: &str| {
        value
            .trim()
            .replace(',', ".")
            .parse::<f64>()
            .map_err(|error| format!("Invalid PowerPoint {label}: {error}"))
    };
    Ok(PowerPointNativeSelection {
        shape_name: fields[0].to_string(),
        slide_index: fields[1]
            .parse::<u32>()
            .map_err(|error| format!("Invalid PowerPoint slide index: {error}"))?,
        slide_id: fields
            .get(6)
            .map(|value| {
                value
                    .parse::<u32>()
                    .map_err(|error| format!("Invalid PowerPoint slide id: {error}"))
            })
            .transpose()?,
        presentation_identity: fields.get(7).map(|value| (*value).to_string()),
        left: parse_number(fields[2], "left")?,
        top: parse_number(fields[3], "top")?,
        width: parse_number(fields[4], "width")?,
        height: parse_number(fields[5], "height")?,
    })
}

fn selection_script(rename_to: Option<&str>) -> String {
    let rename = rename_to
        .map(|value| format!("set name of sr to \"{value}\"\n"))
        .unwrap_or_default();
    format!(
        r#"tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "No active PowerPoint presentation"
set sel to selection of active window
if selection type of sel is not selection type shapes then error "Select exactly one VisualTeX formula shape"
set sr to shape range of sel
if (count of shapes of sr) is not 1 then error "Select exactly one VisualTeX formula shape"
{rename}set currentSlide to slide of view of active window
set presentationIdentity to name of active presentation as text
try
set presentationPath to full name of active presentation as text
if presentationPath is not "" then set presentationIdentity to presentationPath
end try
set separator to ASCII character 31
return (name of sr as text) & separator & (slide index of currentSlide as text) & separator & (left position of sr as text) & separator & (top of sr as text) & separator & (width of sr as text) & separator & (height of sr as text) & separator & (slide id of currentSlide as text) & separator & presentationIdentity
end tell"#
    )
}

pub fn selected_shape() -> Result<PowerPointNativeSelection, String> {
    parse_selection(&run_applescript_with_timeout(
        &selection_script(None),
        APPLESCRIPT_QUERY_TIMEOUT,
    )?)
}

pub fn mark_selected_formula(formula_id: &str) -> Result<PowerPointNativeSelection, String> {
    if !valid_formula_id(formula_id) {
        return Err("Invalid VisualTeX formula id".to_string());
    }
    let shape_name = format!("{SHAPE_PREFIX}{formula_id}");
    parse_selection(&run_applescript_with_timeout(
        &selection_script(Some(&shape_name)),
        APPLESCRIPT_QUERY_TIMEOUT,
    )?)
}

fn replacement_render_height_ratio(
    render_height_px: f64,
    previous_render_height_px: Option<f64>,
) -> Result<Option<f64>, String> {
    if !render_height_px.is_finite() || render_height_px <= 0.0 {
        return Err("PowerPoint formula export has an invalid render height".to_string());
    }
    previous_render_height_px
        .map(|previous_height| {
            if !previous_height.is_finite() || previous_height <= 0.0 {
                return Err(
                    "PowerPoint formula metadata has an invalid previous render height"
                        .to_string(),
                );
            }
            Ok(render_height_px / previous_height)
        })
        .transpose()
}

pub fn upsert_formula_picture_from_clipboard(
    formula_id: &str,
    svg_path: &str,
    width: f64,
    height: f64,
    render_height_px: f64,
    previous_render_height_px: Option<f64>,
    replace_existing: bool,
    original_slide_index: Option<u32>,
    original_shape_name: Option<&str>,
    expected_presentation_identity: Option<&str>,
    target_slide_id: Option<u32>,
    target_slide_index: Option<u32>,
) -> Result<PowerPointNativeSelection, String> {
    if !valid_formula_id(formula_id)
        || ![width, height].into_iter().all(f64::is_finite)
        || width <= 0.0
        || height <= 0.0
    {
        return Err("Invalid VisualTeX native picture insertion request".to_string());
    }
    let render_height_ratio =
        replacement_render_height_ratio(render_height_px, previous_render_height_px)?;
    let svg_path = applescript_string(svg_path)?;
    let expected_presentation_identity = expected_presentation_identity
        .map(applescript_string)
        .transpose()?
        .unwrap_or_else(|| "missing value".to_string());
    let shape_name = format!("{SHAPE_PREFIX}{formula_id}");
    let attempt_id = uuid::Uuid::new_v4().simple().to_string();
    let pending_name = format!("VisualTeXPending_{attempt_id}");
    let original_name = format!("VisualTeXOriginal_{attempt_id}");
    if target_slide_index.is_some_and(|index| index == 0)
        || original_slide_index.is_some_and(|index| index == 0)
        || original_slide_index.is_some() != original_shape_name.is_some()
    {
        return Err("Invalid VisualTeX target slide index".to_string());
    }
    let replacement_geometry = render_height_ratio
        .map(|height_ratio| {
            format!(
                r#"set originalLeft to left position of originalShape
set originalTop to top of originalShape
set originalWidth to width of originalShape
set originalHeight to height of originalShape
set targetHeight to originalHeight * {height_ratio}
set targetWidth to targetHeight * ({width} / {height})
set targetLeft to originalLeft + ((originalWidth - targetWidth) / 2)
set targetTop to originalTop + ((originalHeight - targetHeight) / 2)
set targetRotation to rotation of originalShape"#,
            )
        })
        .unwrap_or_else(|| {
            format!(
                r#"set originalLeft to left position of originalShape
set originalTop to top of originalShape
set originalWidth to width of originalShape
set originalHeight to height of originalShape
set targetHeight to originalHeight
set targetWidth to targetHeight * ({width} / {height})
set targetLeft to originalLeft + ((originalWidth - targetWidth) / 2)
set targetTop to originalTop
set targetRotation to rotation of originalShape"#,
            )
        });
    let replacement_setup = if replace_existing {
        if let (Some(slide_index), Some(original_shape_name)) =
            (original_slide_index, original_shape_name)
        {
            let original_shape_name = applescript_string(original_shape_name)?;
            format!(
                r#"set currentSlide to slide {slide_index} of active presentation
if not (exists shape {original_shape_name} of currentSlide) then error "The selected VisualTeX formula no longer exists"
set originalShape to shape {original_shape_name} of currentSlide
{replacement_geometry}"#
            )
        } else {
            format!(
                r#"set matchingSlideIndexes to {{}}
repeat with candidateSlideNumber from 1 to count of slides of active presentation
set candidateSlide to slide candidateSlideNumber of active presentation
if exists shape "{shape_name}" of candidateSlide then set end of matchingSlideIndexes to slide index of candidateSlide
end repeat
if (count of matchingSlideIndexes) is not 1 then error "The selected VisualTeX formula is not unique in the active presentation"
set currentSlide to slide (item 1 of matchingSlideIndexes) of active presentation
set originalShape to shape "{shape_name}" of currentSlide
{replacement_geometry}"#
            )
        }
    } else {
        let slide_selection = if let Some(slide_id) = target_slide_id {
            format!(
                r#"set matchingTargetSlides to {{}}
repeat with candidateSlideNumber from 1 to count of slides of active presentation
set candidateSlide to slide candidateSlideNumber of active presentation
if slide id of candidateSlide is {slide_id} then set end of matchingTargetSlides to candidateSlide
end repeat
if (count of matchingTargetSlides) is not 1 then error "The original PowerPoint slide no longer exists"
set currentSlide to item 1 of matchingTargetSlides"#
            )
        } else {
            target_slide_index
                .map(|index| format!("set currentSlide to slide {index} of active presentation"))
                .unwrap_or_else(|| "set currentSlide to slide of view of active window".to_string())
        };
        format!(
            r#"{slide_selection}
set targetWidth to {width}
set targetHeight to {height}
set slideWidth to slide width of page setup of active presentation
set targetLeft to (slideWidth - targetWidth) / 2
set targetTop to 72
set targetRotation to 0"#
        )
    };
    let cleanup_conflicting_shape =
        if replace_existing && original_shape_name.is_some_and(|name| name != shape_name) {
            format!(
                r#"repeat with cleanupSlideNumber from 1 to count of slides of active presentation
set cleanupSlide to slide cleanupSlideNumber of active presentation
if exists shape "{shape_name}" of cleanupSlide then delete shape "{shape_name}" of cleanupSlide
end repeat"#
            )
        } else {
            String::new()
        };
    let replacement_finish = if replace_existing {
        format!(
            r#"{cleanup_conflicting_shape}
set name of originalShape to "{original_name}"
set originalShape to shape "{original_name}" of currentSlide
set name of insertedShape to "{shape_name}"
set insertedShape to shape "{shape_name}" of currentSlide"#
        )
    } else {
        format!(
            r#"set name of insertedShape to "{shape_name}"
set insertedShape to shape "{shape_name}" of currentSlide"#
        )
    };
    let rollback = if replace_existing {
        format!(
            r#"try
if exists shape "{original_name}" of currentSlide then
if exists shape "{shape_name}" of currentSlide then delete shape "{shape_name}" of currentSlide
if exists shape "{pending_name}" of currentSlide then delete shape "{pending_name}" of currentSlide
set name of shape "{original_name}" of currentSlide to "{shape_name}"
else
if exists shape "{pending_name}" of currentSlide then delete shape "{pending_name}" of currentSlide
end if
end try"#
        )
    } else {
        format!(
            r#"try
if exists shape "{pending_name}" of currentSlide then delete shape "{pending_name}" of currentSlide
if insertedPromoted then
if exists shape "{shape_name}" of currentSlide then delete shape "{shape_name}" of currentSlide
end if
end try"#
        )
    };
    let script = format!(
        r#"use framework "AppKit"
use framework "Foundation"
use scripting additions

set clipboardObject to current application's NSPasteboard's generalPasteboard()
set savedClipboardItems to current application's NSMutableArray's array()
set sourceItems to clipboardObject's pasteboardItems()
if sourceItems is not missing value then
repeat with sourceItem in sourceItems
set copiedItem to current application's NSPasteboardItem's alloc()'s init()
repeat with sourceType in (sourceItem's types())
set sourceData to sourceItem's dataForType:sourceType
if sourceData is not missing value then copiedItem's setData:sourceData forType:sourceType
end repeat
savedClipboardItems's addObject:copiedItem
end repeat
end if

set svgData to current application's NSData's dataWithContentsOfFile:{svg_path}
if svgData is missing value then error "Unable to load the formula SVG"
clipboardObject's clearContents()
set clipboardWritten to clipboardObject's setData:svgData forType:"com.microsoft.image-svg-xml"
if not clipboardWritten then
clipboardObject's clearContents()
clipboardObject's writeObjects:savedClipboardItems
error "Unable to prepare the PowerPoint SVG clipboard"
end if
set ownedClipboardChangeCount to (clipboardObject's changeCount()) as integer
set insertedPromoted to false

try
with timeout of 12 seconds
tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "No active PowerPoint presentation"
if {expected_presentation_identity} is not missing value then
set activePresentationIdentity to name of active presentation as text
try
set activePresentationPath to full name of active presentation as text
if activePresentationPath is not "" then set activePresentationIdentity to activePresentationPath
end try
if activePresentationIdentity is not {expected_presentation_identity} then error "The active PowerPoint presentation changed while the formula editor was open"
end if
{replacement_setup}
go to slide (view of active window) number (slide index of currentSlide)
set shapeCountBefore to count of shapes of currentSlide
if ((clipboardObject's changeCount()) as integer) is not ownedClipboardChangeCount then error "The clipboard changed before PowerPoint could paste the formula"
paste object (view of active window)
set shapeCountAfter to count of shapes of currentSlide
if shapeCountAfter is not shapeCountBefore + 1 then error "Pasting the formula did not create exactly one PowerPoint shape"

-- Put the newly pasted picture at its final geometry immediately. On edits it
-- exactly overlaps the old picture, so the user never sees a natural-size SVG
-- flash at PowerPoint's default paste position while the host normalizes it.
set insertedShape to shape shapeCountAfter of currentSlide
set left position of insertedShape to targetLeft
set top of insertedShape to targetTop
set lock aspect ratio of insertedShape to false
set width of insertedShape to targetWidth
set height of insertedShape to targetHeight
set lock aspect ratio of insertedShape to true
set rotation of insertedShape to targetRotation

-- PowerPoint initially exposes a temporary SVG paste object, then replaces or
-- normalizes it asynchronously. Properties assigned immediately after `paste
-- object` can therefore be reset a fraction of a second later (for example,
-- VisualTeX_<id> becomes "Graphic 3" and its natural size returns). Let that
-- conversion settle before applying the durable identity and geometry.
delay 0.12
if (count of shapes of currentSlide) is not shapeCountAfter then error "PowerPoint changed the slide while finalizing the formula"
set insertedShape to shape shapeCountAfter of currentSlide
set name of insertedShape to "{pending_name}"
set insertedShape to shape "{pending_name}" of currentSlide
set left position of insertedShape to targetLeft
set top of insertedShape to targetTop
set lock aspect ratio of insertedShape to false
set width of insertedShape to targetWidth
set height of insertedShape to targetHeight
set lock aspect ratio of insertedShape to true
set rotation of insertedShape to targetRotation
{replacement_finish}
set insertedPromoted to true

-- Re-assert the final properties while the host finishes its last SVG
-- normalization passes. Re-resolving the last shape also survives a host-side
-- replacement that invalidates the original AppleScript shape reference.
repeat with finalizeAttempt from 1 to 2
delay 0.08
if exists shape "{shape_name}" of currentSlide then
set insertedShape to shape "{shape_name}" of currentSlide
else
set insertedShape to shape (count of shapes of currentSlide) of currentSlide
set name of insertedShape to "{shape_name}"
set insertedShape to shape "{shape_name}" of currentSlide
end if
set left position of insertedShape to targetLeft
set top of insertedShape to targetTop
set lock aspect ratio of insertedShape to false
set width of insertedShape to targetWidth
set height of insertedShape to targetHeight
set lock aspect ratio of insertedShape to true
set rotation of insertedShape to targetRotation
end repeat
if not (exists shape "{shape_name}" of currentSlide) then error "PowerPoint did not retain the VisualTeX formula identity"
set insertedShape to shape "{shape_name}" of currentSlide
-- Do not rely on PowerPoint preserving the selection created by `paste object`.
-- The Office.js finalizer uses this explicit selection only as a compatibility
-- fallback; its primary lookup is the immutable slide/name/geometry payload.
select insertedShape
set separator to ASCII character 31
set resultPayload to (name of insertedShape as text) & separator & (slide index of currentSlide as text) & separator & (left position of insertedShape as text) & separator & (top of insertedShape as text) & separator & (width of insertedShape as text) & separator & (height of insertedShape as text)
-- Restore the user's clipboard before the old shape is deleted. This keeps
-- clipboard restoration inside the rollback-safe part of the transaction.
if ((clipboardObject's changeCount()) as integer) is ownedClipboardChangeCount then
clipboardObject's clearContents()
clipboardObject's writeObjects:savedClipboardItems
end if
-- Deleting the original is the transaction's final host mutation. Every
-- operation that can reject the replacement has already completed, so an
-- earlier error still reaches the rollback block with the old shape intact.
if {replace_existing} then delete originalShape
end tell
end timeout
if ((clipboardObject's changeCount()) as integer) is ownedClipboardChangeCount then
clipboardObject's clearContents()
clipboardObject's writeObjects:savedClipboardItems
end if
return resultPayload
on error errorMessage number errorNumber
if ((clipboardObject's changeCount()) as integer) is ownedClipboardChangeCount then
clipboardObject's clearContents()
clipboardObject's writeObjects:savedClipboardItems
end if
try
with timeout of 3 seconds
tell application "Microsoft PowerPoint"
{rollback}
end tell
end timeout
end try
error errorMessage number errorNumber
end try"#
    );
    parse_selection(&run_applescript(&script)?)
}

pub fn active_slide_snapshot() -> Result<PowerPointNativeSlideSnapshot, String> {
    let output = run_applescript_with_timeout(
        r#"tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "No active PowerPoint presentation"
set currentSlide to slide of view of active window
set fieldSeparator to "<VISUALTEX_PPT_FIELD>"
set recordSeparator to "<VISUALTEX_PPT_RECORD>"
set presentationIdentity to name of active presentation as text
try
    set presentationPath to full name of active presentation as text
    if presentationPath is not "" then set presentationIdentity to presentationPath
end try
set shapeNamesPayload to ""
set currentShapeCount to count of shapes of currentSlide
-- Enumerating `repeat with candidateShape in shapes ...` can leave Word/PowerPoint
-- automation waiting on a transient SVG proxy object. Resolve each durable shape
-- by index instead; the current count is captured once so the snapshot is atomic.
repeat with shapeIndex from 1 to currentShapeCount
    set candidateName to name of shape shapeIndex of currentSlide as text
    if shapeNamesPayload is "" then
        set shapeNamesPayload to candidateName
    else
        set shapeNamesPayload to shapeNamesPayload & recordSeparator & candidateName
    end if
end repeat
return presentationIdentity & fieldSeparator & (slide index of currentSlide as text) & fieldSeparator & (slide id of currentSlide as text) & fieldSeparator & (currentShapeCount as text) & fieldSeparator & shapeNamesPayload
end tell"#,
        APPLESCRIPT_QUERY_TIMEOUT,
    )?;
    parse_slide_snapshot(&output)
}

fn parse_slide_snapshot(output: &str) -> Result<PowerPointNativeSlideSnapshot, String> {
    let fields: Vec<&str> = output
        .split(POWERPOINT_SNAPSHOT_FIELD_SEPARATOR)
        .collect();
    if fields.len() != 5 {
        return Err(format!(
            "PowerPoint returned an invalid slide snapshot payload: {output}"
        ));
    }
    let shape_names = if fields[4].is_empty() {
        Vec::new()
    } else {
        fields[4]
            .split(POWERPOINT_SNAPSHOT_RECORD_SEPARATOR)
            .map(str::to_string)
            .collect()
    };
    Ok(PowerPointNativeSlideSnapshot {
        presentation_identity: fields[0].to_string(),
        slide_index: fields[1]
            .parse::<u32>()
            .map_err(|error| format!("Invalid PowerPoint slide index: {error}"))?,
        slide_id: fields[2]
            .parse::<u32>()
            .map_err(|error| format!("Invalid PowerPoint slide id: {error}"))?,
        shape_count: fields[3]
            .parse::<u32>()
            .map_err(|error| format!("Invalid PowerPoint shape count: {error}"))?,
        shape_names,
    })
}

fn applescript_string(value: &str) -> Result<String, String> {
    if value.len() > 512 || value.chars().any(char::is_control) {
        return Err("PowerPoint shape name contains unsupported characters".to_string());
    }
    Ok(format!(
        "\"{}\"",
        value.replace('\\', "\\\\").replace('"', "\\\"")
    ))
}

fn applescript_name_list(shape_names: &[String]) -> Result<String, String> {
    if shape_names.len() > 4096 {
        return Err("PowerPoint slide contains too many shapes".to_string());
    }
    shape_names
        .iter()
        .map(|name| applescript_string(name))
        .collect::<Result<Vec<_>, _>>()
        .map(|items| format!("{{{}}}", items.join(", ")))
}

pub fn mark_last_inserted_formula(
    formula_id: &str,
    previous_shape_names: &[String],
) -> Result<PowerPointNativeSelection, String> {
    if !valid_formula_id(formula_id) {
        return Err("Invalid VisualTeX formula id".to_string());
    }
    let shape_name = format!("{SHAPE_PREFIX}{formula_id}");
    let before_names = applescript_name_list(previous_shape_names)?;
    let script = format!(
        r#"tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "No active PowerPoint presentation"
set currentSlide to slide of view of active window
set beforeNames to {before_names}
set insertedShapes to {{}}
repeat with candidateShape in shapes of currentSlide
    set candidateName to name of candidateShape as text
    if beforeNames does not contain candidateName then set end of insertedShapes to candidateShape
end repeat
if (count of insertedShapes) is 1 then
    set targetShape to item 1 of insertedShapes
else
    set sel to selection of active window
    if selection type of sel is not selection type shapes then error "PowerPoint did not expose the inserted formula shape"
    set targetShape to shape range of sel
    if (count of shapes of targetShape) is not 1 then error "PowerPoint returned an ambiguous inserted formula selection"
end if
set name of targetShape to "{shape_name}"
set separator to ASCII character 31
return (name of targetShape as text) & separator & (slide index of currentSlide as text) & separator & (left position of targetShape as text) & separator & (top of targetShape as text) & separator & (width of targetShape as text) & separator & (height of targetShape as text)
end tell"#
    );
    parse_selection(&run_applescript(&script)?)
}

pub fn replace_last_inserted_formula(
    formula_id: &str,
    previous_shape_names: &[String],
    original_shape_name: &str,
    left: f64,
    top: f64,
    width: f64,
    height: f64,
) -> Result<PowerPointNativeSelection, String> {
    if !valid_formula_id(formula_id)
        || !valid_shape_name(original_shape_name)
        || ![left, top, width, height].into_iter().all(f64::is_finite)
        || width <= 0.0
        || height <= 0.0
    {
        return Err("Invalid VisualTeX native replacement reference".to_string());
    }
    let before_names = applescript_name_list(previous_shape_names)?;
    let shape_name = format!("{SHAPE_PREFIX}{formula_id}");
    let temporary_name = format!("VisualTeXOld_{formula_id}");
    let script = format!(
        r#"tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "No active PowerPoint presentation"
set currentSlide to slide of view of active window
if not (exists shape "{original_shape_name}" of currentSlide) then error "The original VisualTeX formula shape no longer exists"
set beforeNames to {before_names}
set insertedShapes to {{}}
repeat with candidateShape in shapes of currentSlide
    set candidateName to name of candidateShape as text
    if beforeNames does not contain candidateName then set end of insertedShapes to candidateShape
end repeat
set originalShape to shape "{original_shape_name}" of currentSlide
if (count of insertedShapes) is 0 then
    set targetShape to originalShape
else if (count of insertedShapes) is 1 then
    set targetShape to item 1 of insertedShapes
else
    set sel to selection of active window
    if selection type of sel is not selection type shapes then error "PowerPoint exposed multiple possible replacement shapes"
    set targetShape to shape range of sel
    if (count of shapes of targetShape) is not 1 then error "PowerPoint returned an ambiguous replacement selection"
end if
set targetNameBefore to name of targetShape as text
if targetNameBefore is not "{original_shape_name}" then
    set name of originalShape to "{temporary_name}"
    set name of targetShape to "{shape_name}"
    set left position of targetShape to {left}
    set top of targetShape to {top}
    set width of targetShape to {width}
    set height of targetShape to {height}
    delete originalShape
else
    set name of targetShape to "{shape_name}"
    set left position of targetShape to {left}
    set top of targetShape to {top}
    set width of targetShape to {width}
    set height of targetShape to {height}
end if
set separator to ASCII character 31
return (name of targetShape as text) & separator & (slide index of currentSlide as text) & separator & (left position of targetShape as text) & separator & (top of targetShape as text) & separator & (width of targetShape as text) & separator & (height of targetShape as text)
end tell"#
    );
    parse_selection(&run_applescript(&script)?)
}

pub fn delete_shape(slide_index: u32, shape_name: &str) -> Result<(), String> {
    if slide_index == 0 || !valid_shape_name(shape_name) {
        return Err("Invalid PowerPoint native shape reference".to_string());
    }
    let script = format!(
        r#"tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "No active PowerPoint presentation"
set targetSlide to slide {slide_index} of active presentation
if exists shape "{shape_name}" of targetSlide then delete shape "{shape_name}" of targetSlide
end tell"#
    );
    run_applescript(&script).map(|_| ())
}

pub fn formula_id_from_shape_name(shape_name: &str) -> Option<String> {
    let formula_id = shape_name.strip_prefix(SHAPE_PREFIX)?;
    valid_formula_id(formula_id).then(|| formula_id.to_string())
}

fn valid_formula_id(value: &str) -> bool {
    uuid::Uuid::parse_str(value)
        .map(|id| id.get_version_num() == 4)
        .unwrap_or(false)
}

fn valid_shape_name(value: &str) -> bool {
    !value.is_empty()
        && value.len() <= 128
        && value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b' '))
}

fn word_formula_after_double_click<ReadSelection, Wait>(
    mut read_selection: ReadSelection,
    mut wait: Wait,
) -> Option<WordNativeFormulaSelection>
where
    ReadSelection: FnMut() -> Result<WordNativeFormulaSelection, String>,
    Wait: FnMut(Duration),
{
    // Word also commits an inline-picture selection after the second mouse
    // down. Retry briefly so a fast double-click cannot be lost while Word is
    // still replacing the text insertion caret with the picture selection.
    for delay in [80_u64, 100, 160, 240] {
        wait(Duration::from_millis(delay));
        let Ok(selection) = read_selection() else {
            continue;
        };
        if selection.marker.starts_with(WORD_METADATA_PREFIX)
            && selection.width.is_finite()
            && selection.height.is_finite()
            && selection.width > 0.0
            && selection.height > 0.0
        {
            return Some(selection);
        }
    }
    None
}

fn powerpoint_selection_after_double_click<ReadSelection, Wait>(
    mut read_selection: ReadSelection,
    mut wait: Wait,
) -> Option<(PowerPointNativeSelection, Option<String>)>
where
    ReadSelection: FnMut() -> Result<PowerPointNativeSelection, String>,
    Wait: FnMut(Duration),
{
    // PowerPoint commits its selection asynchronously after the second
    // mouse-down. Preserve the last valid shape even when the host has renamed
    // a pasted SVG to `Graphic N`; the Office.js command page can still verify
    // durable VisualTeX tags and silently ignore an ordinary picture.
    let mut last_selection = None;
    for delay in [120_u64, 120, 220] {
        wait(Duration::from_millis(delay));
        let Ok(selection) = read_selection() else {
            continue;
        };
        if let Some(formula_id) = formula_id_from_shape_name(&selection.shape_name) {
            return Some((selection, Some(formula_id)));
        }
        last_selection = Some(selection);
    }
    last_selection.map(|selection| (selection, None))
}

#[cfg(target_os = "macos")]
pub fn start_double_click_monitor(bus: PowerPointInteractionBus) -> Result<(), String> {
    use block2::RcBlock;
    use objc2_app_kit::{NSEvent, NSEventMask};
    use std::ptr::NonNull;
    use std::sync::atomic::{AtomicBool, Ordering};

    static STARTED: AtomicBool = AtomicBool::new(false);
    if STARTED.swap(true, Ordering::AcqRel) {
        return Ok(());
    }

    let handler = RcBlock::new(move |event: NonNull<NSEvent>| {
        let event = unsafe { event.as_ref() };
        if event.clickCount() != 2 {
            return;
        }
        let frontmost = frontmost_bundle_id();
        let bus = bus.clone();
        if frontmost.as_deref() == Some(POWERPOINT_BUNDLE_ID) {
            let native_plugin_loaded = native_offline_plugin_loaded("powerpoint");
            std::thread::spawn(move || {
                let Some((selection, formula_id)) =
                    powerpoint_selection_after_double_click(selected_shape, std::thread::sleep)
                else {
                    return;
                };
                if native_plugin_loaded {
                    if formula_id.is_some() {
                        let _ = crate::office::macos_offline::run_double_click_edit_macro(
                            crate::office::sessions::OfficeHost::Powerpoint,
                        );
                    }
                    return;
                }
                if let Some(formula_id) = formula_id {
                    bus.push_powerpoint_edit_selected(selection, formula_id);
                } else {
                    bus.push_powerpoint_edit_requested(selection);
                }
            });
        } else if frontmost.as_deref() == Some(WORD_BUNDLE_ID) {
            let native_plugin_loaded = native_offline_plugin_loaded("word");
            std::thread::spawn(move || {
                let Some(selection) =
                    word_formula_after_double_click(selected_word_formula, std::thread::sleep)
                else {
                    return;
                };
                if !selection.marker.starts_with(WORD_METADATA_PREFIX) {
                    return;
                }
                if native_plugin_loaded {
                    let _ = crate::office::macos_offline::run_double_click_edit_macro(
                        crate::office::sessions::OfficeHost::Word,
                    );
                    return;
                }
                bus.push_word_edit_selected(selection);
            });
        }
    });
    let monitor = NSEvent::addGlobalMonitorForEventsMatchingMask_handler(
        NSEventMask::LeftMouseDown,
        &handler,
    )
    .ok_or_else(|| "macOS did not create the Office double-click monitor".to_string())?;

    // AppKit owns the monitor until it is explicitly removed. VisualTeX keeps it for
    // the lifetime of the background process, so intentionally retain the token.
    std::mem::forget(monitor);
    Ok(())
}

#[cfg(not(target_os = "macos"))]
pub fn start_double_click_monitor(_bus: PowerPointInteractionBus) -> Result<(), String> {
    Ok(())
}

#[cfg(target_os = "macos")]
fn native_offline_plugin_loaded(host: &str) -> bool {
    let file_name = match host {
        "word" => "word.json",
        "powerpoint" => "powerpoint.json",
        _ => return false,
    };
    let home = std::env::var_os("HOME").map(PathBuf::from);
    let Some(home) = home else {
        return false;
    };
    let path = match host {
        "word" => home.join(
            "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime/OfficePluginStatus",
        ),
        "powerpoint" => home.join(
            "Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXRuntime/OfficePluginStatus",
        ),
        _ => return false,
    }
    .join(file_name);
    let Ok(bytes) = fs::read(path) else {
        return false;
    };
    let Ok(value) = serde_json::from_slice::<serde_json::Value>(&bytes) else {
        return false;
    };
    value.get("loaded").and_then(serde_json::Value::as_bool) == Some(true)
        && value.get("host").and_then(serde_json::Value::as_str) == Some(host)
        && value
            .get("pluginVersion")
            .and_then(serde_json::Value::as_str)
            == Some(env!("CARGO_PKG_VERSION"))
}

#[cfg(target_os = "macos")]
fn frontmost_bundle_id() -> Option<String> {
    use objc2_app_kit::NSWorkspace;

    NSWorkspace::sharedWorkspace()
        .frontmostApplication()
        .and_then(|application| application.bundleIdentifier())
        .map(|identifier| identifier.to_string())
}

#[cfg(target_os = "macos")]
fn selected_word_formula() -> Result<WordNativeFormulaSelection, String> {
    let output = run_applescript_with_timeout(
        r#"tell application "Microsoft Word"
if not (exists active document) then error "No active Word document"
set sel to selection
if (count of inline shapes of sel) is 1 then
    set formulaPicture to inline shape 1 of sel
else
    set selectionStart to selection start of sel
    set selectionEnd to selection end of sel
    set probeStart to selectionStart
    if probeStart is greater than 0 then set probeStart to probeStart - 1
    set probeEnd to selectionEnd + 1
    set probeRange to create range active document start probeStart end probeEnd
    if (count of inline shapes of probeRange) is not 1 then error "Select exactly one VisualTeX formula picture"
    set formulaPicture to inline shape 1 of probeRange
end if
set fieldSeparator to "<VISUALTEX_WORD_FIELD>"
return ((alternative text of formulaPicture) as text) & fieldSeparator & (width of formulaPicture as text) & fieldSeparator & (height of formulaPicture as text)
end tell"#,
        APPLESCRIPT_QUERY_TIMEOUT,
    )?;
    let fields = output
        .split(WORD_SELECTION_FIELD_SEPARATOR)
        .collect::<Vec<_>>();
    if fields.len() != 3 {
        return Err(format!("Word returned an invalid formula selection payload: {output}"));
    }
    let parse_number = |value: &str, label: &str| {
        value
            .trim()
            .replace(',', ".")
            .parse::<f64>()
            .map_err(|error| format!("Invalid Word formula {label}: {error}"))
    };
    Ok(WordNativeFormulaSelection {
        marker: fields[0].to_string(),
        width: parse_number(fields[1], "width")?,
        height: parse_number(fields[2], "height")?,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn selection_payload_round_trips() {
        let parsed = parse_selection("VisualTeX_123\u{1f}2\u{1f}10.5\u{1f}20\u{1f}30\u{1f}40")
            .expect("selection payload");
        assert_eq!(parsed.slide_index, 2);
        assert_eq!(parsed.slide_id, None);
        assert_eq!(parsed.presentation_identity, None);
        assert_eq!(parsed.left, 10.5);
        assert_eq!(parsed.height, 40.0);

        let enriched = parse_selection(
            "VisualTeX_123\u{1f}2\u{1f}10.5\u{1f}20\u{1f}30\u{1f}40\u{1f}256\u{1f}/tmp/example.pptx",
        )
        .expect("enriched selection payload");
        assert_eq!(enriched.slide_id, Some(256));
        assert_eq!(
            enriched.presentation_identity.as_deref(),
            Some("/tmp/example.pptx")
        );
    }

    #[test]
    fn slide_snapshot_keeps_every_shape_name() {
        let parsed = parse_slide_snapshot(&format!(
            "/tmp/example.pptx{field}2{field}512{field}3{field}Title 1{record}Graphic 3{record}VisualTeX_00000000-0000-4000-8000-000000000001",
            field = POWERPOINT_SNAPSHOT_FIELD_SEPARATOR,
            record = POWERPOINT_SNAPSHOT_RECORD_SEPARATOR,
        ))
        .expect("slide snapshot payload");
        assert_eq!(parsed.presentation_identity, "/tmp/example.pptx");
        assert_eq!(parsed.slide_index, 2);
        assert_eq!(parsed.slide_id, 512);
        assert_eq!(parsed.shape_count, 3);
        assert_eq!(
            parsed.shape_names,
            vec![
                "Title 1",
                "Graphic 3",
                "VisualTeX_00000000-0000-4000-8000-000000000001",
            ]
        );

        let empty = parse_slide_snapshot(&format!(
            "Untitled{field}1{field}256{field}0{field}",
            field = POWERPOINT_SNAPSHOT_FIELD_SEPARATOR,
        ))
        .expect("empty slide snapshot payload");
        assert!(empty.shape_names.is_empty());
    }

    #[test]
    fn event_bus_is_cursor_based_and_bounded() {
        let bus = PowerPointInteractionBus::default();
        bus.push_edit_selected(
            "powerpoint",
            "VisualTeX_00000000-0000-4000-8000-000000000001".to_string(),
            "00000000-0000-4000-8000-000000000001".to_string(),
        );
        let events = bus.take_after("powerpoint", 0);
        assert_eq!(events.len(), 1);
        assert_eq!(events[0].host, "powerpoint");
        assert!(bus.take_after("word", 0).is_empty());
        assert!(bus.take_after("powerpoint", 0).is_empty());
        assert!(bus.take_after("powerpoint", events[0].cursor).is_empty());
    }

    #[test]
    fn powerpoint_event_keeps_the_double_clicked_shape_snapshot() {
        let bus = PowerPointInteractionBus::default();
        let formula_id = "00000000-0000-4000-8000-000000000001";
        bus.push_powerpoint_edit_selected(
            PowerPointNativeSelection {
                shape_name: format!("VisualTeX_{formula_id}"),
                slide_index: 2,
                slide_id: Some(512),
                presentation_identity: Some("/tmp/example.pptx".to_string()),
                left: 10.0,
                top: 20.0,
                width: 30.0,
                height: 40.0,
            },
            formula_id.to_string(),
        );

        let event = bus.take_after("powerpoint", 0).remove(0);
        assert_eq!(event.slide_index, Some(2));
        assert_eq!(event.slide_id, Some(512));
        assert_eq!(
            event.presentation_identity.as_deref(),
            Some("/tmp/example.pptx")
        );
        assert_eq!(event.left, Some(10.0));
        assert_eq!(event.height, Some(40.0));
    }

    #[test]
    fn event_delivery_is_isolated_per_office_host() {
        let bus = PowerPointInteractionBus::default();
        bus.push_edit_selected(
            "powerpoint",
            "VisualTeX_00000000-0000-4000-8000-000000000001".to_string(),
            "00000000-0000-4000-8000-000000000001".to_string(),
        );
        bus.push_edit_selected("word", "VisualTeX Word Formula".to_string(), String::new());

        let word_events = bus.take_after("word", 0);
        assert_eq!(word_events.len(), 1);
        assert_eq!(word_events[0].host, "word");
        let powerpoint_events = bus.take_after("powerpoint", 0);
        assert_eq!(powerpoint_events.len(), 1);
        assert_eq!(powerpoint_events[0].host, "powerpoint");
    }

    #[test]
    fn word_double_click_retries_and_preserves_marker_and_geometry() {
        let marker = "visualtex:v1:deflate:AbC_123-def";
        let mut attempts = 0;
        let mut waits = Vec::new();
        let resolved = word_formula_after_double_click(
            || {
                attempts += 1;
                if attempts < 3 {
                    return Err("Word selection is still a caret".to_string());
                }
                Ok(WordNativeFormulaSelection {
                    marker: marker.to_string(),
                    width: 84.0,
                    height: 21.0,
                })
            },
            |delay| waits.push(delay),
        )
        .expect("Word formula selection");

        assert_eq!(attempts, 3);
        assert_eq!(waits, vec![Duration::from_millis(80), Duration::from_millis(100), Duration::from_millis(160)]);
        assert_eq!(resolved.marker, marker);
        assert_eq!(resolved.width, 84.0);
        assert_eq!(resolved.height, 21.0);
    }

    #[test]
    fn double_click_selection_retries_until_powerpoint_updates_the_shape() {
        let formula_id = "00000000-0000-4000-8000-000000000001";
        let mut attempts = 0;
        let mut waits = Vec::new();
        let resolved = powerpoint_selection_after_double_click(
            || {
                attempts += 1;
                Ok(PowerPointNativeSelection {
                    shape_name: if attempts < 3 {
                        "Picture 1".to_string()
                    } else {
                        format!("VisualTeX_{formula_id}")
                    },
                    slide_index: 1,
                    slide_id: Some(256),
                    presentation_identity: Some("/tmp/example.pptx".to_string()),
                    left: 0.0,
                    top: 0.0,
                    width: 100.0,
                    height: 30.0,
                })
            },
            |delay| waits.push(delay),
        )
        .expect("formula selection");

        assert_eq!(attempts, 3);
        assert_eq!(waits.len(), 3);
        assert_eq!(resolved.1.as_deref(), Some(formula_id));
    }

    #[test]
    fn powerpoint_replacement_keeps_the_previous_visual_scale() {
        assert_eq!(replacement_render_height_ratio(40.0, Some(40.0)), Ok(Some(1.0)));
        assert_eq!(replacement_render_height_ratio(80.0, Some(40.0)), Ok(Some(2.0)));
        assert_eq!(replacement_render_height_ratio(40.0, None), Ok(None));
        assert!(replacement_render_height_ratio(0.0, Some(40.0)).is_err());
        assert!(replacement_render_height_ratio(40.0, Some(0.0)).is_err());

        // A length-only edit has the same natural render height. The native
        // geometry therefore keeps the old shape height while width grows with
        // the new aspect ratio, so glyphs do not become progressively smaller.
        let original_height = 20.0;
        let height_ratio = replacement_render_height_ratio(40.0, Some(40.0))
            .unwrap()
            .unwrap();
        let first_width = original_height * height_ratio * (400.0 / 40.0);
        let second_width = original_height * height_ratio * (600.0 / 40.0);
        assert_eq!(original_height * height_ratio, 20.0);
        assert_eq!(first_width, 200.0);
        assert_eq!(second_width, 300.0);
    }

    #[test]
    fn shape_formula_id_is_strict() {
        let id = "00000000-0000-4000-8000-000000000001";
        assert_eq!(
            formula_id_from_shape_name(&format!("VisualTeX_{id}")),
            Some(id.to_string())
        );
        assert!(formula_id_from_shape_name("Picture 3").is_none());
    }
}
