use serde::{Deserialize, Serialize};
use std::collections::VecDeque;
use std::sync::{Arc, Mutex};
use std::time::{SystemTime, UNIX_EPOCH};

const SHAPE_PREFIX: &str = "VisualTeX_";
const MAX_EVENTS: usize = 64;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct PowerPointNativeSelection {
    pub shape_name: String,
    pub slide_index: u32,
    pub left: f64,
    pub top: f64,
    pub width: f64,
    pub height: f64,
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
    pub created_at: u64,
}

#[derive(Debug, Default)]
struct InteractionState {
    next_cursor: u64,
    delivered_cursor: u64,
    events: VecDeque<PowerPointInteractionEvent>,
}

#[derive(Debug, Clone, Default)]
pub struct PowerPointInteractionBus {
    inner: Arc<Mutex<InteractionState>>,
}

impl PowerPointInteractionBus {
    pub fn push_edit_selected(&self, host: &'static str, shape_name: String, formula_id: String) {
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
                kind: "edit-selected",
                formula_id,
                shape_name,
                created_at,
            });
            while state.events.len() > MAX_EVENTS {
                state.events.pop_front();
            }
        }
    }

    pub fn take_after(&self, cursor: u64) -> Vec<PowerPointInteractionEvent> {
        self.inner
            .lock()
            .map(|mut state| {
                let threshold = cursor.max(state.delivered_cursor);
                let events = state
                    .events
                    .iter()
                    .filter(|event| event.cursor > threshold)
                    .cloned()
                    .collect::<Vec<_>>();
                if let Some(last) = events.last() {
                    state.delivered_cursor = last.cursor;
                }
                events
            })
            .unwrap_or_default()
    }
}

fn unavailable() -> String {
    "macOS AppleScript PowerPoint integration is unavailable on this platform".to_string()
}

pub fn selected_shape() -> Result<PowerPointNativeSelection, String> {
    Err(unavailable())
}

pub fn mark_selected_formula(_formula_id: &str) -> Result<PowerPointNativeSelection, String> {
    Err(unavailable())
}

#[allow(clippy::too_many_arguments)]
pub fn upsert_formula_picture_from_clipboard(
    _formula_id: &str,
    _svg_path: &str,
    _width: f64,
    _height: f64,
    _replace_existing: bool,
    _original_slide_index: Option<u32>,
    _original_shape_name: Option<&str>,
    _expected_presentation_identity: Option<&str>,
    _target_slide_id: Option<u32>,
    _target_slide_index: Option<u32>,
) -> Result<PowerPointNativeSelection, String> {
    Err(unavailable())
}

pub fn active_slide_snapshot() -> Result<PowerPointNativeSlideSnapshot, String> {
    Err(unavailable())
}

pub fn mark_last_inserted_formula(
    _formula_id: &str,
    _previous_shape_names: &[String],
) -> Result<PowerPointNativeSelection, String> {
    Err(unavailable())
}

#[allow(clippy::too_many_arguments)]
pub fn replace_last_inserted_formula(
    _formula_id: &str,
    _previous_shape_names: &[String],
    _original_shape_name: &str,
    _left: f64,
    _top: f64,
    _width: f64,
    _height: f64,
) -> Result<PowerPointNativeSelection, String> {
    Err(unavailable())
}

pub fn delete_shape(_slide_index: u32, _shape_name: &str) -> Result<(), String> {
    Err(unavailable())
}

pub fn formula_id_from_shape_name(shape_name: &str) -> Option<String> {
    let formula_id = shape_name.strip_prefix(SHAPE_PREFIX)?;
    uuid::Uuid::parse_str(formula_id)
        .ok()
        .filter(|id| id.get_version_num() == 4)
        .map(|_| formula_id.to_string())
}

pub fn start_double_click_monitor(_bus: PowerPointInteractionBus) -> Result<(), String> {
    Ok(())
}
