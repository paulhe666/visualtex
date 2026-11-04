//! Agent writes on the Mac (docs/office-mcp/MACOS.md). As on Windows, each tool puts its target in place
//! through the add-in and runs the Ribbon's command there (`hosts::dispatch`). A command that opens the
//! formula editor hands its request back instead; the app draws the agent's LaTeX in the hidden editor,
//! which applies it as Insert does (macos_offline::open_agent_session). Document operations (redraw,
//! restore, import) run in their own VisualTeX window; the tool waits until it has written.
use super::hosts::{self, CommandError, HostInfo};
use crate::office::macos_offline;
use serde_json::{json, Map, Value};
use std::collections::HashSet;
use std::time::{Duration, Instant};

const STEP_DEADLINE: Duration = Duration::from_secs(120);
const WRITE_DEADLINE: Duration = Duration::from_secs(600);
const FORMULA_KINDS: [&str; 2] = ["picture", "office_equation"];

fn invalid(message: impl Into<String>) -> CommandError {
    CommandError::new("invalid_arguments", message)
}

fn unsupported(message: impl Into<String>) -> CommandError {
    CommandError::new("unsupported_in_host", message)
}

fn text<'a>(arguments: &'a Map<String, Value>, name: &str) -> Option<&'a str> {
    arguments.get(name).and_then(Value::as_str).filter(|value| !value.trim().is_empty())
}

fn required<'a>(arguments: &'a Map<String, Value>, name: &str) -> Result<&'a str, CommandError> {
    text(arguments, name).ok_or_else(|| invalid(format!("{name} is required")))
}

fn number(arguments: &Map<String, Value>, name: &str) -> Option<f64> {
    arguments.get(name).and_then(Value::as_f64)
}

fn format_of(arguments: &Map<String, Value>, name: &str) -> Result<Option<String>, CommandError> {
    match text(arguments, name) {
        None => Ok(None),
        Some(kind) if FORMULA_KINDS.contains(&kind) => Ok(Some(kind.to_string())),
        Some(kind) => Err(invalid(format!("Unknown {name} {kind}: use picture or office_equation"))),
    }
}

/// The formula's rows: `lines`, or `latex` as one row.
fn lines_of(arguments: &Map<String, Value>) -> Result<Option<Vec<String>>, CommandError> {
    if let Some(lines) = arguments.get("lines").filter(|value| !value.is_null()) {
        let rows: Vec<String> = lines
            .as_array()
            .ok_or_else(|| invalid("lines must be an array of strings"))?
            .iter()
            .map(|row| row.as_str().map(str::to_string).ok_or_else(|| invalid("lines must be an array of strings")))
            .collect::<Result<_, _>>()?;
        if rows.is_empty() || rows.iter().all(|row| row.trim().is_empty()) {
            return Err(invalid("lines is empty"));
        }
        return Ok(Some(rows));
    }
    Ok(text(arguments, "latex").map(|latex| vec![latex.trim().to_string()]))
}

fn font_size(arguments: &Map<String, Value>) -> Result<Option<f64>, CommandError> {
    match number(arguments, "font_size_pt") {
        None => Ok(None),
        Some(size) if (5.0..=200.0).contains(&size) => Ok(Some(size)),
        Some(_) => Err(invalid("font_size_pt must be from 5 to 200")),
    }
}

/// One write at a time in each application: the editor and the document belong to one operation.
fn host_lock(host_id: &str) -> &'static tokio::sync::Mutex<()> {
    static WORD: tokio::sync::Mutex<()> = tokio::sync::Mutex::const_new(());
    static POWERPOINT: tokio::sync::Mutex<()> = tokio::sync::Mutex::const_new(());
    static WPS_WRITER: tokio::sync::Mutex<()> = tokio::sync::Mutex::const_new(());
    static WPS_PRESENTATION: tokio::sync::Mutex<()> = tokio::sync::Mutex::const_new(());
    match host_id {
        "p" => &POWERPOINT,
        "ww" => &WPS_WRITER,
        "wp" => &WPS_PRESENTATION,
        _ => &WORD,
    }
}

/// What the add-in's write left: `{message?, formula?, ...}`, with `message` joined from the steps.
struct Outcome {
    messages: Vec<String>,
    data: Map<String, Value>,
}

impl Outcome {
    fn new() -> Self {
        Self { messages: Vec::new(), data: Map::new() }
    }

    fn note(&mut self, data: &Value) {
        if let Some(message) = data.get("message").and_then(Value::as_str).filter(|message| !message.is_empty()) {
            self.messages.push(message.to_string());
        }
    }

    fn say(&mut self, message: impl Into<String>) {
        self.messages.push(message.into());
    }

    fn warn(&mut self, warnings: Vec<String>) {
        if warnings.is_empty() {
            return;
        }
        let list = self.data.entry("warnings").or_insert_with(|| json!([]));
        if let Some(list) = list.as_array_mut() {
            list.extend(warnings.into_iter().map(Value::String));
        }
    }

    fn finish(mut self) -> Value {
        let message = if self.messages.is_empty() { Value::Null } else { Value::String(self.messages.join(" ")) };
        self.data.insert("message".into(), message);
        self.data.entry("warnings").or_insert_with(|| json!([]));
        Value::Object(self.data)
    }
}

struct Writer<'a> {
    host: &'a HostInfo,
    key: String,
}

impl Writer<'_> {
    fn powerpoint(&self) -> bool {
        self.host.app == "powerpoint"
    }

    async fn call(&self, operation: &str, mut arguments: Value, deadline: Duration) -> Result<Value, CommandError> {
        arguments["document"] = Value::String(self.key.clone());
        hosts::dispatch(&self.host.host_id, operation, arguments, deadline).await
    }

    /// Waits for a session the add-in's command started (a formula drawn in the hidden editor, or a
    /// document operation's window) until it has written; what its window would have shown the user
    /// (the formulas it skipped and why).
    async fn wait(&self, session_id: &str, deadline: Duration) -> Result<Vec<String>, CommandError> {
        let started = Instant::now();
        loop {
            if let Some(outcome) = macos_offline::take_agent_session_outcome(session_id) {
                let notes = macos_offline::take_agent_session_notes(session_id);
                return outcome.map(|()| notes).map_err(|error| CommandError::new("operation_failed", error));
            }
            if started.elapsed() > deadline {
                macos_offline::forget_agent_session(session_id);
                return Err(CommandError::new("host_unresponsive", "VisualTeX did not finish writing in time"));
            }
            tokio::time::sleep(Duration::from_millis(100)).await;
        }
    }

    /// The formula session a command prepared, drawn with `lines` (none: a conversion keeps the formula's).
    async fn draw(&self, begun: &Value, lines: Option<Vec<String>>, font_size_pt: Option<f64>) -> Result<Vec<String>, CommandError> {
        let Some(request) = begun.get("request").filter(|request| request.is_object()).cloned() else {
            return Err(CommandError::new("operation_failed", "The add-in did not prepare the formula"));
        };
        let app = hosts::app().ok_or_else(|| CommandError::new("internal", "VisualTeX is still starting"))?;
        let session_id = tokio::task::spawn_blocking(move || macos_offline::open_agent_session(&app, request, lines, font_size_pt))
            .await
            .map_err(|error| CommandError::new("internal", error.to_string()))?
            .map_err(|error| CommandError::new("operation_failed", error))?;
        self.wait(&session_id, WRITE_DEADLINE).await
    }

    /// A document operation the command started in its own window, if it needed one.
    async fn follow(&self, begun: &Value) -> Result<Vec<String>, CommandError> {
        match begun.get("session_id").and_then(Value::as_str) {
            Some(session_id) => {
                macos_offline::watch_agent_session(session_id);
                self.wait(session_id, WRITE_DEADLINE).await
            }
            None => Ok(Vec::new()),
        }
    }

    /// A formula VisualTeX wrote, by the id its request gave it.
    async fn visualtex_formula(&self, vtx_id: &str) -> Result<Option<Value>, CommandError> {
        let found = self.call("find_vtx_formula", json!({ "vtxId": vtx_id }), STEP_DEADLINE).await?;
        Ok(found.get("formula").filter(|formula| formula.is_object()).cloned())
    }

    /// The document's native equations, all pages.
    async fn equations(&self) -> Result<Vec<Value>, CommandError> {
        self.all_formulas(&["office_equation"]).await
    }

    async fn all_formulas(&self, kinds: &[&str]) -> Result<Vec<Value>, CommandError> {
        let mut all = Vec::new();
        loop {
            let page = self
                .call("list_formulas", json!({ "offset": all.len(), "limit": 200, "contextChars": 0, "kinds": kinds }), WRITE_DEADLINE)
                .await?;
            let items = page.get("formulas").and_then(Value::as_array).cloned().unwrap_or_default();
            let total = page.get("total").and_then(Value::as_u64).unwrap_or(0) as usize;
            let empty = items.is_empty();
            all.extend(items);
            if empty || all.len() >= total {
                return Ok(all);
            }
        }
    }

    async fn read(&self, formula_id: &str) -> Result<Value, CommandError> {
        self.call("read_formula", json!({ "formulaId": formula_id }), STEP_DEADLINE).await
    }

    /// The formula ids on a slide (get_slide_layout).
    async fn slide_formulas(&self, slide: Option<u64>) -> Result<(u64, HashSet<String>), CommandError> {
        let mut arguments = json!({});
        if let Some(slide) = slide {
            arguments["slide"] = json!(slide);
        }
        let layout = self.call("get_slide_layout", arguments, STEP_DEADLINE).await?;
        let ids = layout
            .get("shapes")
            .and_then(Value::as_array)
            .into_iter()
            .flatten()
            .filter_map(|shape| shape.get("formula_ids").and_then(Value::as_array))
            .flatten()
            .filter_map(Value::as_str)
            .map(str::to_string)
            .collect();
        Ok((layout.get("slide").and_then(Value::as_u64).unwrap_or(1), ids))
    }

    async fn set_frame(&self, formula_id: &str, position: Value) -> Result<Value, CommandError> {
        self.call("set_frame", json!({ "formulaId": formula_id, "position": position }), STEP_DEADLINE).await
    }

    // insert_formula

    async fn insert(&self, arguments: &Map<String, Value>) -> Result<Value, CommandError> {
        let lines = lines_of(arguments)?.ok_or_else(|| invalid("latex is required"))?;
        let display = match text(arguments, "display") {
            None | Some("inline") => "inline",
            Some("block") => "block",
            Some(other) => return Err(invalid(format!("Unknown display {other}: use inline or block"))),
        };
        let format = format_of(arguments, "format")?.unwrap_or_else(|| "picture".to_string());
        let numbered = arguments.get("numbered").and_then(Value::as_bool).unwrap_or(false);
        if numbered && (display != "block" || self.powerpoint()) {
            return Err(invalid("Only display formulas in Word are numbered"));
        }
        let location = arguments.get("location").cloned().unwrap_or(Value::Null);
        let placed = location.get("left_pt").is_some() || location.get("top_pt").is_some();
        let slide = location.get("slide").and_then(Value::as_u64);
        if self.powerpoint() && placed && slide.is_none() {
            return Err(invalid("A position on a slide needs location.slide"));
        }
        let native_on_slide = self.powerpoint() && format == "office_equation";
        let before = if native_on_slide { Some(self.slide_formulas(slide).await?) } else { None };
        // a Word equation is found by its LaTeX among the equations before and after
        let word_equations = if !self.powerpoint() && format == "office_equation" {
            Some(self.equations().await?)
        } else {
            None
        };
        // WPS Writer's request names no formula id (the editor gives the picture one): the new picture is
        // the one that was not there before
        let wps_pictures = if !self.powerpoint() && format == "picture" && self.host.vendor == "kingsoft" {
            Some(self.all_formulas(&["picture"]).await?)
        } else {
            None
        };
        let begun = self
            .call("begin_insert", json!({ "display": display, "format": format, "numbered": numbered, "location": location }), STEP_DEADLINE)
            .await?;
        self.draw(&begun, Some(lines), font_size(arguments)?).await?;

        let mut outcome = Outcome::new();
        outcome.note(&begun);
        let mut formula = match (&before, begun.pointer("/request/formulaId").and_then(Value::as_str)) {
            (None, Some(vtx_id)) => self.visualtex_formula(vtx_id).await?,
            (Some((slide, ids)), _) => {
                let (_, after) = self.slide_formulas(Some(*slide)).await?;
                let added: Vec<&String> = after.difference(ids).collect();
                match added.as_slice() {
                    [id] => Some(self.read(id).await?),
                    _ => None,
                }
            }
            _ => None,
        };
        if let (None, Some(earlier)) = (&formula, &wps_pictures) {
            let known: HashSet<&str> = earlier.iter().filter_map(|item| item.get("formula_id").and_then(Value::as_str)).collect();
            formula = self
                .all_formulas(&["picture"])
                .await?
                .into_iter()
                .find(|item| item.get("formula_id").and_then(Value::as_str).is_some_and(|id| !known.contains(id)));
        }
        if let (None, Some(earlier)) = (&formula, &word_equations) {
            let now = self.equations().await?;
            formula = now
                .iter()
                .find(|candidate| {
                    let latex = candidate.get("latex");
                    now.iter().filter(|other| other.get("latex") == latex).count()
                        > earlier.iter().filter(|other| other.get("latex") == latex).count()
                })
                .cloned();
        }
        if placed {
            if let Some(id) = formula.as_ref().and_then(|formula| formula.get("formula_id")).and_then(Value::as_str).map(str::to_string) {
                let mut position = json!({});
                for field in ["left_pt", "top_pt"] {
                    if let Some(value) = location.get(field) {
                        position[field] = value.clone();
                    }
                }
                self.set_frame(&id, position).await?;
                formula = Some(self.read(&id).await?);
            } else {
                outcome.say("The formula is on the slide; list the slide's formulas to move it.");
            }
        }
        outcome.say("Inserted the formula.");
        if let Some(formula) = formula {
            outcome.data.insert("formula".into(), formula);
        }
        Ok(outcome.finish())
    }

    // update_formula

    async fn update(&self, arguments: &Map<String, Value>) -> Result<Value, CommandError> {
        let formula_id = required(arguments, "formula_id")?.to_string();
        let lines = lines_of(arguments)?;
        let format = format_of(arguments, "format")?;
        let size = font_size(arguments)?;
        let position = arguments.get("position").filter(|value| !value.is_null()).cloned();
        if position.as_ref().is_some_and(|position| !position.is_object()) {
            return Err(invalid("position must be an object"));
        }
        if lines.is_none() && format.is_none() && position.is_none() && size.is_none() {
            return Err(invalid("Give latex (new content), format (convert), font_size_pt or position"));
        }
        let current = self.read(&formula_id).await?;
        let kind = current.get("kind").and_then(Value::as_str).unwrap_or_default().to_string();
        let converts = format.as_deref().is_some_and(|format| format != kind);

        if self.powerpoint() && converts {
            return self.replace_on_slide(&current, format.as_deref().unwrap_or("picture"), lines, size, position).await;
        }
        let mut outcome = Outcome::new();
        let mut formula_id = formula_id;
        let mut formula = None;
        if lines.is_some() || converts {
            let begun = self
                .call("begin_update", json!({ "formulaId": formula_id, "format": format }), STEP_DEADLINE)
                .await?;
            outcome.note(&begun);
            if begun.get("session_id").and_then(Value::as_str).is_some() {
                self.draw(&begun, lines, size).await?;
            }
            formula = match begun.pointer("/request/formulaId").and_then(Value::as_str) {
                Some(vtx_id) => self.visualtex_formula(vtx_id).await?,
                None => None,
            };
            if formula.is_none() {
                formula = self.selected_formula().await?;
            }
            if let Some(id) = formula.as_ref().and_then(|formula| formula.get("formula_id")).and_then(Value::as_str) {
                formula_id = id.to_string();
            }
        } else if let Some(size) = size {
            let set = self.call("set_formula_font_size", json!({ "formulaId": formula_id, "font_size_pt": size }), STEP_DEADLINE).await?;
            outcome.note(&set);
        }
        if let Some(position) = position {
            if !self.powerpoint() {
                return Err(invalid("position is for formulas on slides (PowerPoint)"));
            }
            self.set_frame(&formula_id, position).await?;
            formula = Some(self.read(&formula_id).await?);
        }
        if formula.is_none() {
            formula = self.read(&formula_id).await.ok();
        }
        outcome.say("Updated the formula.");
        if let Some(formula) = formula {
            outcome.data.insert("formula".into(), formula);
        }
        Ok(outcome.finish())
    }

    async fn selected_formula(&self) -> Result<Option<Value>, CommandError> {
        let selection = self.call("get_selection", json!({}), STEP_DEADLINE).await?;
        Ok(selection.get("formula").filter(|formula| formula.is_object()).cloned())
    }

    /// PowerPoint has no conversion of one formula: the formula is drawn anew in the other format where it
    /// was, and the old one deleted.
    async fn replace_on_slide(
        &self,
        current: &Value,
        format: &str,
        lines: Option<Vec<String>>,
        size: Option<f64>,
        position: Option<Value>,
    ) -> Result<Value, CommandError> {
        let old_id = current.get("formula_id").and_then(Value::as_str).unwrap_or_default().to_string();
        let lines = match lines {
            Some(lines) => lines,
            None => current
                .get("lines")
                .and_then(Value::as_array)
                .map(|rows| rows.iter().filter_map(Value::as_str).map(str::to_string).collect::<Vec<_>>())
                .filter(|rows| !rows.is_empty())
                .or_else(|| current.get("latex").and_then(Value::as_str).map(|latex| vec![latex.to_string()]))
                .ok_or_else(|| CommandError::new("operation_failed", "The formula's LaTeX could not be read"))?,
        };
        let slide = current.pointer("/location/slide").and_then(Value::as_u64).unwrap_or(1);
        let frame = current.get("frame").cloned().unwrap_or(Value::Null);
        let mut arguments = Map::new();
        arguments.insert("lines".into(), json!(lines));
        arguments.insert("format".into(), json!(format));
        arguments.insert("display".into(), current.get("display").cloned().unwrap_or(json!("block")));
        let size = size.or_else(|| current.get("font_size_pt").and_then(Value::as_f64));
        if let Some(size) = size {
            arguments.insert("font_size_pt".into(), json!(size.clamp(5.0, 200.0)));
        }
        let mut location = json!({ "slide": slide });
        for field in ["left_pt", "top_pt"] {
            if let Some(value) = position.as_ref().and_then(|position| position.get(field)).or_else(|| frame.get(field)) {
                location[field] = value.clone();
            }
        }
        arguments.insert("location".into(), location);
        let inserted = self.insert(&arguments).await?;
        self.call("delete_formula", json!({ "formulaId": old_id }), STEP_DEADLINE).await?;
        let mut outcome = Outcome::new();
        outcome.say(format!("Drew the formula again as {format} where it was."));
        let mut formula = inserted.get("formula").cloned();
        if let (Some(position), Some(id)) = (
            position.filter(|position| position.get("width_pt").is_some() || position.get("height_pt").is_some()),
            formula.as_ref().and_then(|formula| formula.get("formula_id")).and_then(Value::as_str).map(str::to_string),
        ) {
            self.set_frame(&id, position).await?;
            formula = Some(self.read(&id).await?);
        }
        if let Some(formula) = formula {
            outcome.data.insert("formula".into(), formula);
        }
        Ok(outcome.finish())
    }

    // convert_formulas

    async fn convert(&self, arguments: &Map<String, Value>) -> Result<Value, CommandError> {
        let to = format_of(arguments, "to")?.ok_or_else(|| invalid("to is required"))?;
        let from = required(arguments, "from")?.to_string();
        if !self.powerpoint() {
            if from == "any" {
                return Err(invalid("In Word, from is picture or office_equation"));
            }
            let begun = self.call("convert_formulas", Value::Object(arguments.clone()), STEP_DEADLINE).await?;
            let skipped = self.follow(&begun).await?;
            let mut outcome = Outcome::new();
            outcome.note(&begun);
            outcome.say("Converted the formulas.");
            outcome.warn(skipped);
            return Ok(outcome.finish());
        }
        if from != "any" && !FORMULA_KINDS.contains(&from.as_str()) {
            return Err(invalid(format!("Unknown from {from}")));
        }
        // the formulas in scope, last first (a native equation's neighbours keep their ids)
        let scope = text(arguments, "scope").unwrap_or("document");
        let mut targets: Vec<Value> = match scope {
            "selection" => self.selected_formula().await?.into_iter().collect(),
            "document" | "slide" => {
                let slide = if scope == "slide" {
                    let selection = self.call("get_selection", json!({}), STEP_DEADLINE).await?;
                    selection.pointer("/location/slide").and_then(Value::as_u64)
                } else {
                    None
                };
                self.all_formulas(&[])
                    .await?
                    .into_iter()
                    .filter(|formula| slide.is_none() || formula.pointer("/location/slide").and_then(Value::as_u64) == slide)
                    .collect()
            }
            other => return Err(invalid(format!("Unknown scope {other}"))),
        };
        targets.retain(|formula| {
            let kind = formula.get("kind").and_then(Value::as_str).unwrap_or_default();
            kind != to && (from == "any" || kind == from)
        });
        let mut converted = 0;
        let mut failures = Vec::new();
        for formula in targets.iter().rev() {
            let id = formula.get("formula_id").and_then(Value::as_str).unwrap_or_default();
            let current = match self.read(id).await {
                Ok(current) => current,
                Err(error) => {
                    failures.push(format!("{id}: {}", error.message));
                    continue;
                }
            };
            match self.replace_on_slide(&current, &to, None, None, None).await {
                Ok(_) => converted += 1,
                Err(error) => failures.push(format!("{id}: {}", error.message)),
            }
        }
        let mut outcome = Outcome::new();
        outcome.say(format!("Converted {converted} formula(s) to {to}."));
        if !failures.is_empty() {
            outcome.data.insert("warnings".into(), json!(failures));
        }
        Ok(outcome.finish())
    }

    // latex_text_to_formulas, formulas_to_latex_text

    async fn document_operation(&self, tool: &str, arguments: &Map<String, Value>) -> Result<Value, CommandError> {
        if tool == "latex_text_to_formulas" {
            format_of(arguments, "format")?.ok_or_else(|| invalid("format is required"))?;
        } else {
            format_of(arguments, "from")?.ok_or_else(|| invalid("from is required"))?;
        }
        let begun = self.call(tool, Value::Object(arguments.clone()), STEP_DEADLINE).await?;
        let skipped = self.follow(&begun).await?;
        let mut outcome = Outcome::new();
        outcome.note(&begun);
        outcome.warn(skipped);
        if tool == "latex_text_to_formulas" && arguments.get("number_display_formulas").and_then(Value::as_bool) == Some(true) {
            if self.powerpoint() {
                outcome.say("Slides have no equation numbers.");
            } else {
                let numbered = self
                    .call("number_display_formulas", json!({ "scope": text(arguments, "scope").unwrap_or("document") }), WRITE_DEADLINE)
                    .await?;
                outcome.note(&numbered);
            }
        }
        if tool == "formulas_to_latex_text" && self.powerpoint() {
            outcome.say("PowerPoint turns every formula in the scope into LaTeX.");
        }
        outcome.say("Done.");
        Ok(outcome.finish())
    }

    // insert_rich_text

    async fn rich_text(&self, arguments: &Map<String, Value>) -> Result<Value, CommandError> {
        let text = required(arguments, "text")?.to_string();
        let source_format = match text_or(arguments, "source_format", "auto") {
            "markdown" | "latex" | "auto" => text_or(arguments, "source_format", "auto").to_string(),
            other => return Err(invalid(format!("Unknown source_format {other}"))),
        };
        let output = match format_of(arguments, "format")?.as_deref() {
            Some("picture") => "image",
            _ => "omml",
        };
        let begun = self
            .call("begin_import", json!({ "location": arguments.get("location").cloned().unwrap_or(Value::Null) }), STEP_DEADLINE)
            .await?;
        let request = begun
            .get("request")
            .filter(|request| request.is_object())
            .cloned()
            .ok_or_else(|| CommandError::new("operation_failed", "The add-in did not start the import"))?;
        let agent = json!({
            "text": text,
            "sourceFormat": source_format,
            "outputKind": output,
            "numberDisplayFormulas": arguments.get("number_display_formulas").and_then(Value::as_bool).unwrap_or(false),
        });
        let app = hosts::app().ok_or_else(|| CommandError::new("internal", "VisualTeX is still starting"))?;
        let session_id = tokio::task::spawn_blocking(move || macos_offline::open_agent_import(&app, request, agent))
            .await
            .map_err(|error| CommandError::new("internal", error.to_string()))?
            .map_err(|error| CommandError::new("operation_failed", error))?;
        let skipped = self.wait(&session_id, WRITE_DEADLINE).await?;
        let mut outcome = Outcome::new();
        outcome.say("Inserted the text.");
        outcome.warn(skipped);
        Ok(outcome.finish())
    }
}

fn text_or<'a>(arguments: &'a Map<String, Value>, name: &str, fallback: &'a str) -> &'a str {
    text(arguments, name).unwrap_or(fallback)
}

const WORD_ONLY: [&str; 5] = [
    "insert_rich_text",
    "insert_equation_reference",
    "number_display_formulas",
    "update_equation_numbers",
    "set_equation_number_format",
];

/// One write tool on a document of a running Word or PowerPoint (`key` is the add-in's document key).
pub async fn run(host: &HostInfo, tool: &str, key: String, arguments: Map<String, Value>) -> Result<Value, CommandError> {
    if host.app == "powerpoint" && WORD_ONLY.contains(&tool) {
        return Err(unsupported(format!("{tool} is for Word documents")));
    }
    let _guard = tokio::time::timeout(WRITE_DEADLINE, host_lock(&host.host_id).lock())
        .await
        .map_err(|_| CommandError::new("host_busy", "Another agent write is still running in this application"))?;
    let writer = Writer { host, key };
    match tool {
        "insert_formula" => writer.insert(&arguments).await,
        "update_formula" => writer.update(&arguments).await,
        "convert_formulas" => writer.convert(&arguments).await,
        "latex_text_to_formulas" | "formulas_to_latex_text" => writer.document_operation(tool, &arguments).await,
        "insert_rich_text" => writer.rich_text(&arguments).await,
        "delete_formula" | "select_formula" => {
            let formula_id = required(&arguments, "formula_id")?;
            let data = writer.call(tool, json!({ "formulaId": formula_id }), STEP_DEADLINE).await?;
            Ok(finished(data))
        }
        "set_formula_font_size" => {
            let size = font_size(&arguments)?.ok_or_else(|| invalid("font_size_pt is required"))?;
            if let Some(kind) = text(&arguments, "kind") {
                // every inline or every display formula of a Word document
                if text(&arguments, "formula_id").is_some() {
                    return Err(invalid("Give formula_id (one formula) or kind (every inline or display formula), not both"));
                }
                if writer.powerpoint() {
                    return Err(unsupported("kind sets the formulas of a Word document; on slides, give formula_id"));
                }
                if !matches!(kind, "inline" | "display") {
                    return Err(invalid(format!("Unknown kind {kind}: use inline or display")));
                }
                let data = writer.call(tool, json!({ "kind": kind, "font_size_pt": size }), WRITE_DEADLINE).await?;
                return Ok(finished(data));
            }
            let formula_id = required(&arguments, "formula_id")?;
            let data = writer
                .call(tool, json!({ "formulaId": formula_id, "font_size_pt": size }), STEP_DEADLINE)
                .await?;
            let mut result = finished(data);
            if let Ok(formula) = writer.read(formula_id).await {
                result["formula"] = formula;
            }
            Ok(result)
        }
        "number_display_formulas" | "update_equation_numbers" | "set_equation_number_format" | "insert_equation_reference" => {
            if tool == "insert_equation_reference" {
                required(&arguments, "number")?;
                format_of(&arguments, "source")?;
            }
            let data = writer.call(tool, Value::Object(arguments), WRITE_DEADLINE).await?;
            Ok(finished(data))
        }
        _ => Err(CommandError::new("unknown_tool", format!("Unknown tool {tool}"))),
    }
}

fn finished(data: Value) -> Value {
    let mut outcome = Outcome::new();
    outcome.note(&data);
    if let Some(formula) = data.get("formula").filter(|formula| formula.is_object()) {
        outcome.data.insert("formula".into(), formula.clone());
    }
    for field in ["frame", "frame_of"] {
        if let Some(value) = data.get(field) {
            outcome.data.insert(field.into(), value.clone());
        }
    }
    outcome.finish()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn arguments(value: Value) -> Map<String, Value> {
        value.as_object().cloned().unwrap()
    }

    #[test]
    fn formulas_come_as_latex_or_rows() {
        assert_eq!(lines_of(&arguments(json!({ "latex": " E=mc^2 " }))).unwrap(), Some(vec!["E=mc^2".to_string()]));
        assert_eq!(
            lines_of(&arguments(json!({ "latex": "x", "lines": ["a&=b", "c&=d"] }))).unwrap(),
            Some(vec!["a&=b".to_string(), "c&=d".to_string()])
        );
        assert_eq!(lines_of(&arguments(json!({}))).unwrap(), None);
        assert!(lines_of(&arguments(json!({ "lines": [] }))).is_err());
        assert!(lines_of(&arguments(json!({ "lines": ["a", 3] }))).is_err());
    }

    #[test]
    fn formats_and_sizes_are_checked() {
        assert_eq!(format_of(&arguments(json!({ "format": "picture" })), "format").unwrap().as_deref(), Some("picture"));
        assert_eq!(format_of(&arguments(json!({})), "format").unwrap(), None);
        assert!(format_of(&arguments(json!({ "format": "visualtex_ole" })), "format").is_err());
        assert_eq!(font_size(&arguments(json!({ "font_size_pt": 12 }))).unwrap(), Some(12.0));
        assert!(font_size(&arguments(json!({ "font_size_pt": 400 }))).is_err());
    }

    #[test]
    fn outcomes_join_their_messages() {
        let mut outcome = Outcome::new();
        outcome.note(&json!({ "message": "Updated 3 equation numbers." }));
        outcome.note(&json!({ "message": null }));
        outcome.say("Done.");
        let result = outcome.finish();
        assert_eq!(result["message"], "Updated 3 equation numbers. Done.");
        assert_eq!(result["warnings"], json!([]));
        assert_eq!(finished(json!({ "message": null }))["message"], Value::Null);
    }
}
