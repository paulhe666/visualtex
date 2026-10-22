//! Transient converter jobs. These snapshots never enter the editor SessionStore.
#[cfg(test)]
mod tests {
    use super::*;
    fn request(ids: &[&str]) -> RenderBatchRequest {
        serde_json::from_value(serde_json::json!({
            "id": Uuid::new_v4().to_string(),
            "items": ids.iter().map(|id| serde_json::json!({
                "id": id, "request": {
                    "mode": "create", "host": "word", "codeFormat": "latex",
                    "displayMode": "inline", "objectMode": "wordOmml",
                    "fontSizePt": 12,
                    "lines": [{ "id": Uuid::new_v4().to_string(), "latex": "x^2" }]
                }
            })).collect::<Vec<_>>()
        })).unwrap()
    }
    fn success(id: &str) -> RenderBatchOutput {
        RenderBatchOutput {
            id: id.into(), error: None,
            export: Some(OfficeExportResult {
                svg: String::new(), svg_base64: String::new(),
                math_ml: Some("<math><msup><mi>x</mi><mn>2</mn></msup></math>".into()),
                png_base64: None, width: 20.0, height: 15.0, baseline: Some(12.0),
                formula_letter_font: None, formula_chinese_font: None,
            }),
        }
    }
    #[test]
    fn completion_is_correlated_atomic_and_keeps_input_immutable() {
        let id = create(request(&["a", "b"])).unwrap();
        let original = snapshot(&id).unwrap();
        let receiver = subscribe(&id).unwrap();
        let invalid = RenderBatchCompletion { items: vec![success("a"), success("wrong")] };
        assert!(complete(&id, invalid).is_err());
        assert!(receiver.borrow().is_none());
        complete(&id, RenderBatchCompletion { items: vec![success("b"), success("a")] }).unwrap();
        let result = receiver.borrow().clone().unwrap();
        assert_eq!(result.items[0].id, "a");
        assert_eq!(result.items[0].session.status, OfficeSessionStatus::Committing);
        assert_eq!(result.items[0].session.lines[0].latex, original.items[0].session.lines[0].latex);
        assert!(snapshot(&id).unwrap().items[0].session.export_result.is_none());
        assert!(complete(&id, RenderBatchCompletion { items: vec![success("a"), success("b")] }).is_err());
        release(&id).unwrap();
    }
    #[test]
    fn caller_owned_id_can_be_released_without_receiving_create_response() {
        let request = request(&["a"]);
        let id = request.id.clone();
        assert_eq!(create(request).unwrap(), id);
        let mut duplicate = self::request(&["a"]);
        duplicate.id = id.clone();
        assert!(matches!(create(duplicate), Err(SessionError::Conflict(_))));
        release(&id).unwrap();
        assert!(snapshot(&id).is_err());
    }
    #[test]
    fn release_notifies_pending_waiter_and_is_idempotent() {
        let id = create(request(&["cancel"])).unwrap();
        let receiver = subscribe(&id).unwrap();
        release(&id).unwrap();
        let terminal = receiver.borrow().clone().unwrap();
        assert!(terminal.error.is_some());
        assert_eq!(terminal.items[0].session.status, OfficeSessionStatus::Cancelled);
        assert!(snapshot(&id).is_err());
        release(&id).unwrap();
    }
    #[test]
    fn invalid_item_result_fails_that_item_only_and_converter_failures_are_terminal() {
        let id = create(request(&["good", "nomathml"])).unwrap();
        let receiver = subscribe(&id).unwrap();
        let mut bad = success("nomathml");
        bad.export.as_mut().unwrap().math_ml = None;
        complete(&id, RenderBatchCompletion { items: vec![success("good"), bad] }).unwrap();
        let result = receiver.borrow().clone().unwrap();
        assert_eq!(result.items[0].session.status, OfficeSessionStatus::Committing);
        assert_eq!(result.items[1].session.status, OfficeSessionStatus::Failed);
        assert_eq!(result.items[1].session.error.as_deref(), Some(ITEM_RESULT_MISSING_MATHML));
        release(&id).unwrap();

        let id = create(request(&["a"])).unwrap();
        let receiver = subscribe(&id).unwrap();
        fail(&id, "Invalid Office render batch.").unwrap();
        let result = receiver.borrow().clone().unwrap();
        assert_eq!(result.error_code.as_deref(), Some("converter-failed"));
        assert_eq!(result.error.as_deref(), Some("Invalid Office render batch."));
        release(&id).unwrap();

        let id = create(request(&["a"])).unwrap();
        let receiver = subscribe(&id).unwrap();
        expire_untaken(&id, 20).unwrap();
        assert_eq!(receiver.borrow().clone().unwrap().error_code.as_deref(), Some("converter-not-started"));
        release(&id).unwrap();

        let id = create(request(&["a"])).unwrap();
        let receiver = subscribe(&id).unwrap();
        assert_eq!(take_next().unwrap().as_deref(), Some(id.as_str()));
        expire_untaken(&id, 20).unwrap();
        assert!(receiver.borrow().is_none());
        release(&id).unwrap();
    }
    #[test]
    fn failure_is_an_item_result_and_duplicate_ids_are_rejected() {
        assert!(create(request(&["same", "same"])).is_err());
        let id = create(request(&["good", "bad"])).unwrap();
        let receiver = subscribe(&id).unwrap();
        complete(&id, RenderBatchCompletion { items: vec![
            success("good"), RenderBatchOutput { id: "bad".into(), export: None, error: Some("Unknown command".into()) }
        ] }).unwrap();
        let result = receiver.borrow().clone().unwrap();
        assert_eq!(result.items[1].session.status, OfficeSessionStatus::Failed);
        assert_eq!(result.items[1].session.error.as_deref(), Some("Unknown command"));
        release(&id).unwrap();
    }
}

use super::sessions::{CreateOfficeSessionInput, OfficeExportResult, OfficeFormulaSession,
    OfficeSessionStatus, SessionError, SessionStore};
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet, VecDeque};
use std::sync::{Mutex, OnceLock};
use tokio::sync::watch;
#[cfg(test)]
use uuid::Uuid;


#[derive(Deserialize)]
pub struct RenderBatchRequest { pub id: String, pub items: Vec<RenderBatchInput> }
#[derive(Deserialize)]
pub struct RenderBatchInput { pub id: String, pub request: CreateOfficeSessionInput }
#[derive(Clone, Serialize)]
pub struct RenderBatchItem { pub id: String, pub session: OfficeFormulaSession }
#[derive(Clone, Serialize)]
pub struct RenderBatchSnapshot {
    pub id: String,
    pub items: Vec<RenderBatchItem>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<String>,
    /// why the whole batch failed: "cancelled", "converter-not-started" or
    /// "converter-failed" (the add-ins word their message from it)
    #[serde(rename = "errorCode", skip_serializing_if = "Option::is_none")]
    pub error_code: Option<String>,
}

/// Item failures the add-ins recognise (OfficeErrorText): VisualTeX's own
/// checks of a converter result, worded per formula.
pub const ITEM_RESULT_MISSING_MATHML: &str = "VisualTeX's result for this formula has no MathML.";
pub const ITEM_RESULT_INVALID_SIZE: &str = "VisualTeX's result for this formula has an invalid size (width, height or baseline).";
pub const ITEM_RESULT_MISSING_IMAGE: &str = "VisualTeX's result for this formula has no SVG or PNG image.";
#[derive(Deserialize)]
pub struct RenderBatchCompletion { pub items: Vec<RenderBatchOutput> }
#[derive(Deserialize)]
pub struct RenderBatchOutput {
    pub id: String,
    pub export: Option<OfficeExportResult>,
    pub error: Option<String>,
}
struct Job {
    input: RenderBatchSnapshot,
    terminal: watch::Sender<Option<RenderBatchSnapshot>>,
    /// the converter window took the batch (it is rendering it)
    taken: bool,
}
#[derive(Default)]
struct Registry { jobs: HashMap<String, Job>, pending: VecDeque<String> }
static JOBS: OnceLock<Mutex<Registry>> = OnceLock::new();

fn registry() -> &'static Mutex<Registry> { JOBS.get_or_init(|| Mutex::new(Registry::default())) }
fn unavailable() -> SessionError { SessionError::Io("Render batch registry is unavailable".into()) }

pub fn create(request: RenderBatchRequest) -> Result<String, SessionError> {
    if !super::sessions::valid_uuid(&request.id) {
        return Err(SessionError::Invalid("Render batch ID must be a UUID".into()));
    }
    if request.items.is_empty() {
        return Err(SessionError::Invalid("Render batches require at least one item".into()));
    }
    let mut ids = HashSet::new();
    let mut items = Vec::with_capacity(request.items.len());
    for (index, item) in request.items.into_iter().enumerate() {
        if item.id.trim().is_empty() || !ids.insert(item.id.clone()) {
            return Err(SessionError::Invalid(format!(
                "Formula {} of the render batch has an empty or duplicate item ID", index + 1)));
        }
        // name the formula whose request is invalid (the add-ins show it)
        let mut session = SessionStore::build_render_session(item.request).map_err(|error| match error {
            SessionError::Invalid(message) => SessionError::Invalid(format!("Formula {} ({}): {message}", index + 1, item.id)),
            other => other,
        })?;
        session.auto_commit_on_close = false;
        items.push(RenderBatchItem { id: item.id, session });
    }
    let mut state = registry().lock().map_err(|_| unavailable())?;
    let id = request.id;
    if state.jobs.contains_key(&id) {
        return Err(SessionError::Conflict("Render batch ID is already active".into()));
    }
    let (terminal, _) = watch::channel(None);
    state.jobs.insert(id.clone(), Job {
        input: RenderBatchSnapshot { id: id.clone(), items, error: None, error_code: None }, terminal,
        taken: false,
    });
    state.pending.push_back(id.clone());
    Ok(id)
}
pub fn take_next() -> Result<Option<String>, SessionError> {
    let mut state = registry().lock().map_err(|_| unavailable())?;
    let next = state.pending.pop_front();
    if let Some(job) = next.as_ref().and_then(|id| state.jobs.get_mut(id)) {
        job.taken = true;
    }
    Ok(next)
}

fn finish_with_error(job: &Job, code: &str, message: String) {
    if job.terminal.borrow().is_some() { return; }
    let mut result = job.input.clone();
    result.error = Some(message);
    result.error_code = Some(code.into());
    for item in &mut result.items { item.session.status = OfficeSessionStatus::Failed; }
    job.terminal.send_replace(Some(result));
}

/// The converter window could not render the batch (it reports why).
pub fn fail(id: &str, message: &str) -> Result<(), SessionError> {
    let state = registry().lock().map_err(|_| unavailable())?;
    let job = state.jobs.get(id).ok_or(SessionError::NotFound)?;
    let message = message.trim();
    let message = if message.is_empty() { "The VisualTeX conversion window reported a failure without details" } else { message };
    finish_with_error(job, "converter-failed", message.chars().take(4000).collect());
    Ok(())
}

/// A batch the converter window has not taken after `seconds`: it is not
/// running (or not listening); the waiting add-in is told so at once.
pub fn expire_untaken(id: &str, seconds: u64) -> Result<(), SessionError> {
    let mut state = registry().lock().map_err(|_| unavailable())?;
    let Some(job) = state.jobs.get(id) else { return Ok(()); };
    if job.taken || job.terminal.borrow().is_some() { return Ok(()); }
    finish_with_error(job, "converter-not-started", format!(
        "The VisualTeX conversion window did not start this render batch within {seconds} s"));
    state.pending.retain(|pending| pending != id);
    Ok(())
}
pub fn snapshot(id: &str) -> Result<RenderBatchSnapshot, SessionError> {
    let state = registry().lock().map_err(|_| unavailable())?;
    Ok(state.jobs.get(id).ok_or(SessionError::NotFound)?.input.clone())
}
pub fn subscribe(id: &str) -> Result<watch::Receiver<Option<RenderBatchSnapshot>>, SessionError> {
    let state = registry().lock().map_err(|_| unavailable())?;
    Ok(state.jobs.get(id).ok_or(SessionError::NotFound)?.terminal.subscribe())
}
/// What is wrong with one formula's result, if anything (it fails that formula only).
fn export_problem(export: &OfficeExportResult, mode: &str) -> Option<&'static str> {
    if export.math_ml.as_deref().unwrap_or("").trim().is_empty() {
        return Some(ITEM_RESULT_MISSING_MATHML);
    }
    if !export.width.is_finite() || export.width <= 0.0
        || !export.height.is_finite() || export.height <= 0.0
        || export.baseline.is_some_and(|value| !value.is_finite()) {
        return Some(ITEM_RESULT_INVALID_SIZE);
    }
    if mode != "wordOmml" && mode != "mathTypeOle"
        && (export.svg.trim().is_empty() || export.svg_base64.trim().is_empty()
            || export.png_base64.as_deref().unwrap_or("").trim().is_empty()) {
        return Some(ITEM_RESULT_MISSING_IMAGE);
    }
    None
}
pub fn complete(id: &str, completion: RenderBatchCompletion) -> Result<(), SessionError> {
    let state = registry().lock().map_err(|_| unavailable())?;
    let job = state.jobs.get(id).ok_or(SessionError::NotFound)?;
    if job.terminal.borrow().is_some() {
        return Err(SessionError::Conflict("Render batch is already terminal".into()));
    }
    if completion.items.len() != job.input.items.len() {
        return Err(SessionError::Invalid("Render completion must contain every requested item".into()));
    }
    let mut outputs = HashMap::new();
    for output in completion.items {
        let has_error = output.error.as_deref().is_some_and(|error| !error.trim().is_empty());
        if has_error == output.export.is_some() || outputs.contains_key(&output.id) {
            return Err(SessionError::Invalid("Each render item requires exactly one export or error and a unique ID".into()));
        }
        outputs.insert(output.id.clone(), output);
    }
    let mut result = job.input.clone();
    for item in &mut result.items {
        let output = outputs.remove(&item.id)
            .ok_or_else(|| SessionError::Invalid("Render completion item IDs do not match request".into()))?;
        if let Some(export) = output.export {
            if let Some(problem) = export_problem(&export, &item.session.object_mode) {
                item.session.status = OfficeSessionStatus::Failed;
                item.session.error = Some(problem.into());
                continue;
            }
            item.session.export_width = export.width;
            item.session.export_height = export.height;
            item.session.export_result = Some(export);
            item.session.status = OfficeSessionStatus::Committing;
        } else {
            item.session.status = OfficeSessionStatus::Failed;
            item.session.error = output.error;
        }
    }
    job.terminal.send_replace(Some(result));
    Ok(())
}
pub fn release(id: &str) -> Result<(), SessionError> {
    let mut state = registry().lock().map_err(|_| unavailable())?;
    let Some(job) = state.jobs.remove(id) else { return Ok(()); };
    state.pending.retain(|pending| pending != id);
    if job.terminal.borrow().is_none() {
        let mut result = job.input;
        result.error = Some("Render batch was cancelled".into());
        result.error_code = Some("cancelled".into());
        for item in &mut result.items { item.session.status = OfficeSessionStatus::Cancelled; }
        job.terminal.send_replace(Some(result));
    }
    Ok(())
}
