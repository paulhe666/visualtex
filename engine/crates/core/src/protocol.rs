use std::collections::BTreeMap;
use std::sync::Arc;
use serde::{Deserialize, Serialize};
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::session::{DocumentSnapshot, EngineSession, Revision, SessionSnapshot};

pub const PROTOCOL_VERSION: u32 = 1;
pub const MAX_WIRE_BYTES: usize = 1_048_576;
pub const MAX_SESSIONS: usize = 64;
pub const MAX_PARKED_SESSIONS: usize = 2048;
pub const MAX_PARKED_WEIGHT: usize = 512 * 1024 * 1024;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Request {
    pub protocol: u32,
    pub command: Command,
}

#[derive(Debug, Deserialize)]
#[serde(tag = "op", rename_all = "camelCase", rename_all_fields = "camelCase", deny_unknown_fields)]
pub enum Command {
    Create { document_id: String, source: Option<String> },
    CreateWithSymbols { document_id:String,source:String,symbols:Vec<serde_json::Value> },
    AnalyzeCustomSymbol { definition:serde_json::Value },
    AnalyzeSourceWithSymbols { source:String,symbols:Vec<serde_json::Value> },
    Inspect { handle: u32, expected_revision: Option<Revision> },
    Destroy { handle: u32, expected_revision: Revision },
    ParkSession { handle: u32, expected_revision: Revision },
    ResumeSession { handle: u32, expected_revision: Revision },
    ReplaceSource { handle: u32, expected_revision: Revision, source: String },
    PatchSource { handle: u32, expected_revision: Revision, edit: crate::source_edit::SourceEdit },
    InspectDocument { handle: u32, expected_revision: Option<Revision> },
    InspectEditor { handle: u32, expected_revision: Option<Revision> },
    CopySelection { handle: u32, expected_revision: Revision },
    CaptureOcrTarget { handle: u32, expected_revision: Revision },
    CancelOcrTarget { handle: u32, request_id: u32 },
    EditorAction { handle: u32, expected_revision: Revision, event_sequence: u32, action: crate::editing::EditorAction,
        /// Opt-in: leave `editor.document` out of the reply when the action
        /// did not change the document (same revision and source identity).
        #[serde(default)] omit_unchanged_document: bool },
    /// Stateless syntax diagnostics for source editors and headless validation.
    /// Returned node ids are analysis-local and must not be used as edit anchors.
    AnalyzeSource { source: String },
    Layout { handle: u32, expected_revision: Revision, options: crate::layout::LayoutOptions, include_svg: bool },
    ExportPresentationMathMl { handle: u32, expected_revision: Revision, display_mode: bool },
    QueryScene { handle: u32, expected_revision: Revision, scene_token: u32, query: crate::scene_queries::SceneQuery },
    Stats,
}

#[derive(Debug, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase")]
pub enum Reply {
    Session { handle: u32, snapshot: SessionSnapshot },
    CustomSymbolAnalysis { asset:Arc<crate::custom_symbols::SymbolAsset> },
    Document { handle: u32, snapshot: DocumentSnapshot },
    Editor { handle: u32, snapshot: crate::editing::EditorSnapshot },
    EditorAction { handle: u32, result: crate::editing::EditorActionReply },
    SourcePatched { handle: u32, snapshot: DocumentSnapshot, mapping: crate::source_edit::SourceChangeMap },
    Layout { handle: u32, scene_token: u32, snapshot: Arc<crate::scene::LayoutSnapshot>, svg: Option<String> },
    PresentationMathMl { handle: u32, result: crate::presentation_mathml::PresentationMathMl },
    SceneQuery { handle: u32, scene_token: u32, result: crate::scene_queries::SceneQueryResult },
    FontRegistered { font: crate::fonts::FontIdentity },
    Clipboard { handle: u32, fragment: crate::clipboard::ClipboardFragment },
    OcrTarget { handle: u32, target: crate::ocr::OcrInsertionTarget },
    OcrTargetCancelled { handle: u32, cancelled: bool },
    Destroyed { handle: u32 },
    SessionParked { handle: u32, snapshot: SessionSnapshot },
    SessionResumed { handle: u32, snapshot: SessionSnapshot },
    SourceAnalysis { analysis: crate::parser::ParseResult, lossless: String, canonical: Option<String> },
    Stats { protocol: u32, active_sessions: usize },
}

#[derive(Debug, Serialize)]
#[serde(tag = "status", rename_all = "camelCase")]
pub enum Response {
    Ok { result: Reply },
    Error { error: EngineError },
}

impl Response {
    pub fn json(self) -> String {
        // Geometry is validated as finite before reaching this boundary. Keep the public error
        // bounded even if a future response type introduces a serialization error.
        serde_json::to_string(&self).unwrap_or_else(|_| {
            String::from("{\"status\":\"error\",\"error\":{\"code\":\"invalidInput\",\"message\":\"Response serialization failed\"}}")
        })
    }
}

pub fn error_json(code: ErrorCode, message: &str) -> String {
    Response::Error { error: EngineError::new(code, message) }.json()
}

/// Owns sessions, not pointers into WASM memory. Handles are never recycled.
/// This is shared by the native CLI and the thin wasm-bindgen boundary.
pub struct Runtime {
    sessions: BTreeMap<u32, EngineSession>,
    parked: BTreeMap<u32, EngineSession>,
    parked_weight: usize,
    documents: BTreeMap<String, u32>,
    next_handle: u32,
    fonts: crate::fonts::FontRegistry,
    layouts: BTreeMap<u32, crate::layout::LayoutEngine>,
    scenes: BTreeMap<u32, (u32, Arc<crate::scene::LayoutSnapshot>)>,
    next_scene_token: u32,
}

impl Default for Runtime {
    fn default() -> Self {
        Self { sessions: BTreeMap::new(), parked: BTreeMap::new(), parked_weight: 0,
            documents: BTreeMap::new(), next_handle: 1,
            fonts: crate::fonts::FontRegistry::default(), layouts: BTreeMap::new(),
            scenes: BTreeMap::new(), next_scene_token: 1 }
    }
}

impl Runtime {
    pub fn active_sessions(&self) -> usize { self.sessions.len() }
    pub fn parked_sessions(&self) -> usize { self.parked.len() }
    pub fn parked_weight(&self) -> usize { self.parked_weight }

    fn retained_weight(session: &EngineSession) -> usize {
        let retained_nodes = if session.document.complete { 0 } else {
            let mut count = 0usize;
            if let Some(tree) = &session.last_complete { tree.visit(&mut |_| count += 1); }
            count
        };
        session.document.source.len().saturating_mul(24)
            .saturating_add(session.document.tokens.len().saturating_mul(80))
            .saturating_add((session.document.facts.nodes + retained_nodes).saturating_mul(512))
            .saturating_add(session.document.symbol_assets.retained_bytes())
            .saturating_add(session.editor.history.weight())
    }

    /// Binary resource injection bypasses JSON inflation, not the font budget.
    /// Hosts own file discovery; no font path or URL enters the pure core.
    pub fn register_font(&mut self, bytes: &[u8], face_index: u32) -> EngineResult<crate::fonts::FontIdentity> {
        self.fonts.register(bytes, face_index)
    }

    pub fn register_font_json(&mut self, bytes: &[u8], face_index: u32) -> String {
        match self.register_font(bytes, face_index) {
            Ok(font) => Response::Ok { result: Reply::FontRegistered { font } }.json(),
            Err(error) => Response::Error { error }.json(),
        }
    }

    pub fn dispatch_json(&mut self, input: &str) -> String {
        if input.len() > MAX_WIRE_BYTES {
            return error_json(ErrorCode::ResourceLimit, "Request exceeds the byte budget");
        }
        let request: Request = match serde_json::from_str(input) {
            Ok(request) => request,
            Err(_) => return error_json(ErrorCode::InvalidInput, "Invalid engine request schema"),
        };
        match self.dispatch(request) {
            Ok(result) => Response::Ok { result }.json(),
            Err(error) => Response::Error { error }.json(),
        }
    }

    pub fn dispatch(&mut self, request: Request) -> EngineResult<Reply> {
        if request.protocol != PROTOCOL_VERSION {
            return Err(EngineError::new(ErrorCode::UnsupportedProtocol, "Unsupported engine protocol version"));
        }
        match request.command {
            Command::AnalyzeCustomSymbol{definition}=>Ok(Reply::CustomSymbolAnalysis{asset:crate::custom_symbols::SymbolAsset::freeze(definition)?}),
            Command::AnalyzeSourceWithSymbols{source,symbols}=>{
                let symbols=crate::custom_symbols::SymbolCatalog::freeze(symbols)?;
                let analysis=crate::parser::parse_with_symbols(crate::source::SourceText::new(1,source)?,crate::session::NodeId(1),&mut crate::model::IdAllocator::new(2),symbols)?;
                let lossless=crate::model::latex(&analysis.tree,crate::model::LatexMode::Lossless);
                let canonical=(analysis.complete&&analysis.supported).then(||crate::model::latex(&analysis.tree,crate::model::LatexMode::Canonical));
                Ok(Reply::SourceAnalysis{analysis,lossless,canonical})
            }
            Command::CreateWithSymbols{document_id,source,symbols}=>{
                if self.documents.contains_key(&document_id){return Err(EngineError::new(ErrorCode::Conflict,"Document already has a retained session"));}
                if self.sessions.len()>=MAX_SESSIONS{return Err(EngineError::new(ErrorCode::ResourceLimit,"Active session budget exhausted"));}
                let symbols=crate::custom_symbols::SymbolCatalog::freeze(symbols)?;
                let session=EngineSession::from_source_with_symbols(document_id.clone(),source,symbols)?;
                let next=self.next_handle.checked_add(1).ok_or_else(||EngineError::new(ErrorCode::ResourceLimit,"Session identity space exhausted"))?;
                let handle=self.next_handle;let snapshot=session.snapshot();self.sessions.insert(handle,session);self.documents.insert(document_id,handle);self.next_handle=next;
                Ok(Reply::Session{handle,snapshot})
            }
            Command::Create { document_id, source } => {
                if self.documents.contains_key(&document_id) {
                    return Err(EngineError::new(ErrorCode::Conflict, "Document already has a retained session"));
                }
                if self.sessions.len() >= MAX_SESSIONS {
                    return Err(EngineError::new(ErrorCode::ResourceLimit, "Active session budget exhausted"));
                }
                let session = match source {
                    Some(source) => EngineSession::from_source(document_id.clone(), source)?,
                    None => EngineSession::new(document_id.clone())?,
                };
                let next = self.next_handle.checked_add(1)
                    .ok_or_else(|| EngineError::new(ErrorCode::ResourceLimit, "Session identity space exhausted"))?;
                let handle = self.next_handle;
                let snapshot = session.snapshot();
                self.sessions.insert(handle, session);
                self.documents.insert(document_id, handle);
                self.next_handle = next;
                Ok(Reply::Session { handle, snapshot })
            }
            Command::Inspect { handle, expected_revision } => {
                let session = self.get(handle)?;
                if let Some(expected) = expected_revision { session.check_revision(expected)?; }
                Ok(Reply::Session { handle, snapshot: session.snapshot() })
            }
            Command::Destroy { handle, expected_revision } => {
                if let Some(session) = self.sessions.get(&handle) {
                    session.check_revision(expected_revision)?;
                    let document_id = session.document_id().to_owned();
                    self.sessions.remove(&handle);
                    self.layouts.remove(&handle);
                    self.scenes.remove(&handle);
                    self.documents.remove(&document_id);
                    return Ok(Reply::Destroyed { handle });
                }
                let session = self.parked.get(&handle)
                    .ok_or_else(|| EngineError::new(ErrorCode::InvalidHandle, "Unknown or destroyed session handle"))?;
                session.check_revision(expected_revision)?;
                let document_id = session.document_id().to_owned();
                let weight = Self::retained_weight(session);
                self.parked.remove(&handle);
                self.parked_weight = self.parked_weight.saturating_sub(weight);
                self.documents.remove(&document_id);
                Ok(Reply::Destroyed { handle })
            }
            Command::ParkSession { handle, expected_revision } => {
                let (snapshot, weight) = {
                    let session = self.get(handle)?;
                    session.check_revision(expected_revision)?;
                    if session.editor_snapshot()?.composition.is_some() {
                        return Err(EngineError::new(ErrorCode::Conflict,
                            "Finish or cancel composition before parking the editor session"));
                    }
                    (session.snapshot(), Self::retained_weight(session))
                };
                if self.parked.len() >= MAX_PARKED_SESSIONS
                    || self.parked_weight.saturating_add(weight) > MAX_PARKED_WEIGHT
                {
                    return Err(EngineError::new(ErrorCode::ResourceLimit, "Parked session retention budget exhausted"));
                }
                let mut session = self.sessions.remove(&handle).ok_or_else(||
                    EngineError::new(ErrorCode::InvalidHandle, "Unknown or destroyed session handle"))?;
                session.invalidate_ocr_target();
                self.layouts.remove(&handle);
                self.scenes.remove(&handle);
                self.parked.insert(handle, session);
                self.parked_weight += weight;
                Ok(Reply::SessionParked { handle, snapshot })
            }
            Command::ResumeSession { handle, expected_revision } => {
                if self.sessions.len() >= MAX_SESSIONS {
                    return Err(EngineError::new(ErrorCode::ResourceLimit, "Active session budget exhausted"));
                }
                let (snapshot, weight) = {
                    let session = self.parked.get(&handle)
                        .ok_or_else(|| EngineError::new(ErrorCode::InvalidHandle, "Unknown or already active session handle"))?;
                    session.check_revision(expected_revision)?;
                    (session.snapshot(), Self::retained_weight(session))
                };
                let session = self.parked.remove(&handle).ok_or_else(||
                    EngineError::new(ErrorCode::InvalidHandle, "Unknown or already active session handle"))?;
                self.parked_weight = self.parked_weight.saturating_sub(weight);
                self.sessions.insert(handle, session);
                Ok(Reply::SessionResumed { handle, snapshot })
            }
            Command::AnalyzeSource { source } => {
                let source = crate::source::SourceText::new(1, source)?;
                let analysis = crate::parser::parse(source, crate::session::NodeId(1),
                    &mut crate::model::IdAllocator::new(2))?;
                let lossless = crate::model::latex(&analysis.tree, crate::model::LatexMode::Lossless);
                // Do not offer repaired/opaque input as a normalized replacement.
                let canonical = (analysis.complete && analysis.supported)
                    .then(|| crate::model::latex(&analysis.tree, crate::model::LatexMode::Canonical));
                Ok(Reply::SourceAnalysis { analysis, lossless, canonical })
            }
            Command::ReplaceSource { handle, expected_revision, source } => {
                let session = self.get_mut(handle)?;
                Ok(Reply::Document { handle, snapshot: session.replace_source(expected_revision, source)? })
            }
            Command::PatchSource { handle, expected_revision, edit } => {
                let (snapshot, mapping) = self.get_mut(handle)?.patch_source(expected_revision, edit)?;
                Ok(Reply::SourcePatched { handle, snapshot, mapping })
            }
            Command::InspectDocument { handle, expected_revision } => {
                let session = self.get(handle)?;
                if let Some(expected) = expected_revision { session.check_revision(expected)?; }
                Ok(Reply::Document { handle, snapshot: session.document_snapshot() })
            }
            Command::InspectEditor { handle, expected_revision } => {
                let session = self.get(handle)?;
                if let Some(expected) = expected_revision { session.check_revision(expected)?; }
                Ok(Reply::Editor { handle, snapshot: session.editor_snapshot()? })
            }
            Command::CopySelection { handle, expected_revision } => {
                let fragment = self.get(handle)?.copy_selection(expected_revision)?;
                Ok(Reply::Clipboard { handle, fragment })
            }
            Command::CaptureOcrTarget { handle, expected_revision } => {
                let target=self.get_mut(handle)?.capture_ocr_target(handle,expected_revision)?;
                Ok(Reply::OcrTarget{handle,target})
            }
            Command::CancelOcrTarget { handle, request_id } => {
                let cancelled=self.get_mut(handle)?.cancel_ocr_target(request_id);
                Ok(Reply::OcrTargetCancelled{handle,cancelled})
            }
            Command::EditorAction { handle, expected_revision, event_sequence, action, omit_unchanged_document } => {
                if let crate::editing::EditorAction::ApplyOcrLatex{target,..}=&action {
                    if target.session_handle!=handle{return Err(EngineError::new(ErrorCode::Conflict,"OCR target belongs to another session handle"));}
                }
                let session = self.get_mut(handle)?;
                let before = (session.revision(), session.source_id());
                let result = session.apply_action(expected_revision, event_sequence, action)?;
                let omit_document = omit_unchanged_document && before == (result.editor.document.revision, result.editor.document.source_id);
                Ok(Reply::EditorAction { handle, result: crate::editing::EditorActionReply { result, omit_document } })
            }
            Command::Layout { handle, expected_revision, options, include_svg } => {
                let session = self.sessions.get(&handle)
                    .ok_or_else(|| EngineError::new(ErrorCode::InvalidHandle, "Unknown or destroyed session handle"))?;
                session.check_revision(expected_revision)?;
                let next = self.next_scene_token.checked_add(1)
                    .ok_or_else(|| EngineError::new(ErrorCode::ResourceLimit, "Layout identity space exhausted"))?;
                let snapshot = Arc::new(self.layouts.entry(handle).or_default().layout(session, &self.fonts, &options)?);
                let svg = if include_svg { Some(snapshot.svg()?) } else { None };
                let scene_token = self.next_scene_token;
                self.scenes.insert(handle, (scene_token, Arc::clone(&snapshot)));
                self.next_scene_token = next;
                Ok(Reply::Layout { handle, scene_token, snapshot, svg })
            }
            Command::ExportPresentationMathMl { handle, expected_revision, display_mode } => {
                let session = self.get(handle)?;
                session.check_revision(expected_revision)?;
                let result = crate::presentation_mathml::encode(&session.document, display_mode)?;
                Ok(Reply::PresentationMathMl { handle, result })
            }
            Command::QueryScene { handle, expected_revision, scene_token, query } => {
                let session = self.get(handle)?;
                session.check_revision(expected_revision)?;
                let (token, scene) = self.scenes.get(&handle)
                    .ok_or_else(|| EngineError::new(ErrorCode::Conflict, "Create a layout before querying its geometry"))?;
                if *token != scene_token || scene.revision != expected_revision
                    || scene.source_id != session.document.source.id()
                {
                    return Err(EngineError::new(ErrorCode::Conflict, "Layout token is stale; refresh geometry before applying pointer or selection results"));
                }
                let result = crate::scene_queries::query(scene, &session.document.tree, query)?;
                Ok(Reply::SceneQuery { handle, scene_token, result })
            }
            Command::Stats => Ok(Reply::Stats { protocol: PROTOCOL_VERSION, active_sessions: self.active_sessions() }),
        }
    }

    fn get(&self, handle: u32) -> EngineResult<&EngineSession> {
        self.sessions.get(&handle)
            .ok_or_else(|| EngineError::new(ErrorCode::InvalidHandle, "Unknown or destroyed session handle"))
    }

    fn get_mut(&mut self, handle: u32) -> EngineResult<&mut EngineSession> {
        self.sessions.get_mut(&handle)
            .ok_or_else(|| EngineError::new(ErrorCode::InvalidHandle, "Unknown or destroyed session handle"))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::{Value, json};
    fn send(rt: &mut Runtime, command: Value) -> Value {
        serde_json::from_str(&rt.dispatch_json(&json!({"protocol":1,"command":command}).to_string())).unwrap()
    }
    #[test]
    fn lifecycle_and_stale_requests_are_atomic() {
        let mut rt = Runtime::default();
        assert_eq!(send(&mut rt, json!({"op":"create","documentId":"d"}))["result"]["handle"], 1);
        assert_eq!(send(&mut rt, json!({"op":"create","documentId":"d"}))["error"]["code"], "conflict");
        assert_eq!(send(&mut rt, json!({"op":"destroy","handle":1,"expectedRevision":9}))["error"]["code"], "staleRevision");
        assert_eq!(rt.active_sessions(), 1);
        assert_eq!(send(&mut rt, json!({"op":"destroy","handle":1,"expectedRevision":0}))["status"], "ok");
        assert_eq!(send(&mut rt, json!({"op":"inspect","handle":1}))["error"]["code"], "invalidHandle");
        assert_eq!(send(&mut rt, json!({"op":"create","documentId":"d"}))["result"]["handle"], 2);
    }
    #[test]
    fn unknown_fields_and_bad_numbers_do_not_mutate_state() {
        let mut rt = Runtime::default();
        for input in ["null", "{}", "[]", "not json",
            r#"{"protocol":1,"command":{"op":"create","documentId":"d","unknown":true}}"#,
            r#"{"protocol":1,"command":{"op":"inspect","handle":-1}}"#,
            r#"{"protocol":1,"command":{"op":"inspect","handle":1.2}}"#,
            r#"{"protocol":1,"command":{"op":"stats"},"unexpected":0}"#] {
            let response: Value = serde_json::from_str(&rt.dispatch_json(input)).unwrap();
            assert_eq!(response["error"]["code"], "invalidInput", "{input}");
            assert_eq!(rt.active_sessions(), 0);
        }
    }
    #[test]
    fn protocol_mismatch_and_request_budget() {
        let mut rt = Runtime::default();
        let response: Value = serde_json::from_str(&rt.dispatch_json(r#"{"protocol":2,"command":{"op":"create","documentId":"d"}}"#)).unwrap();
        assert_eq!(response["error"]["code"], "unsupportedProtocol");
        let response: Value = serde_json::from_str(&rt.dispatch_json(&" ".repeat(MAX_WIRE_BYTES + 1))).unwrap();
        assert_eq!(response["error"]["code"], "resourceLimit");
        assert_eq!(rt.active_sessions(), 0);
    }
    #[test]
    fn session_budget_and_released_capacity() {
        let mut rt = Runtime::default();
        for index in 0..MAX_SESSIONS {
            assert_eq!(send(&mut rt, json!({"op":"create","documentId":format!("d{index}")}))["status"], "ok");
        }
        assert_eq!(send(&mut rt, json!({"op":"create","documentId":"overflow"}))["error"]["code"], "resourceLimit");
        send(&mut rt, json!({"op":"destroy","handle":1,"expectedRevision":0}));
        assert_eq!(send(&mut rt, json!({"op":"create","documentId":"replacement"}))["result"]["handle"], 65);
    }
    #[test]
    fn repeated_lifecycle_does_not_leak_or_reuse_handles() {
        let mut rt = Runtime::default();
        for handle in 1..=1000 {
            assert_eq!(send(&mut rt, json!({"op":"create","documentId":"same-document"}))["result"]["handle"], handle);
            send(&mut rt, json!({"op":"destroy","handle":handle,"expectedRevision":0}));
            assert_eq!(rt.active_sessions(), 0);
        }
    }
    #[test]
    fn park_resume_preserves_exact_editor_state_history_and_handle() {
        let mut rt=Runtime::default();
        assert_eq!(send(&mut rt,json!({"op":"create","documentId":"virtual-1","source":"x"}))["result"]["handle"],1);
        assert_eq!(send(&mut rt,json!({"op":"editorAction","handle":1,"expectedRevision":0,"eventSequence":1,
            "action":{"kind":"edit","command":{"kind":"insertText","text":"y","mode":"math"},"groupId":null}}))["status"],"ok");
        assert_eq!(send(&mut rt,json!({"op":"editorAction","handle":1,"expectedRevision":1,"eventSequence":2,
            "action":{"kind":"setSelection","selection":{"kind":"range","anchor":{"kind":"row","nodeId":1,"offset":0,"affinity":"after"},"focus":{"kind":"row","nodeId":1,"offset":0,"affinity":"after"}}}}))["status"],"ok");
        let before=send(&mut rt,json!({"op":"inspectEditor","handle":1,"expectedRevision":1}));
        let parked=send(&mut rt,json!({"op":"parkSession","handle":1,"expectedRevision":1}));
        assert_eq!(parked["result"]["kind"],"sessionParked");assert_eq!(parked["result"]["handle"],1);
        assert_eq!(rt.active_sessions(),0);assert_eq!(rt.parked_sessions(),1);assert!(rt.parked_weight()>0);
        assert_eq!(send(&mut rt,json!({"op":"inspectEditor","handle":1}))["error"]["code"],"invalidHandle");
        assert_eq!(send(&mut rt,json!({"op":"create","documentId":"virtual-1","source":"other"}))["error"]["code"],"conflict");
        let resumed=send(&mut rt,json!({"op":"resumeSession","handle":1,"expectedRevision":1}));
        assert_eq!(resumed["result"]["kind"],"sessionResumed");assert_eq!(resumed["result"]["handle"],1);
        assert_eq!(rt.active_sessions(),1);assert_eq!(rt.parked_sessions(),0);assert_eq!(rt.parked_weight(),0);
        assert_eq!(send(&mut rt,json!({"op":"inspectEditor","handle":1,"expectedRevision":1})),before);
        let undo=send(&mut rt,json!({"op":"editorAction","handle":1,"expectedRevision":1,"eventSequence":3,"action":{"kind":"undo"}}));
        assert_eq!(undo["result"]["result"]["editor"]["document"]["source"],"x");
    }
    #[test]
    fn park_rejects_composition_and_stale_resume_is_atomic() {
        let mut rt=Runtime::default();
        send(&mut rt,json!({"op":"create","documentId":"virtual-ime","source":"x"}));
        assert_eq!(send(&mut rt,json!({"op":"editorAction","handle":1,"expectedRevision":0,"eventSequence":1,
            "action":{"kind":"beginComposition","mode":"math"}}))["status"],"ok");
        assert_eq!(send(&mut rt,json!({"op":"parkSession","handle":1,"expectedRevision":0}))["error"]["code"],"conflict");
        assert_eq!((rt.active_sessions(),rt.parked_sessions()),(1,0));
        assert_eq!(send(&mut rt,json!({"op":"editorAction","handle":1,"expectedRevision":0,"eventSequence":2,
            "action":{"kind":"cancelComposition","id":1}}))["status"],"ok");
        assert_eq!(send(&mut rt,json!({"op":"parkSession","handle":1,"expectedRevision":0}))["status"],"ok");
        let weight=rt.parked_weight();
        assert_eq!(send(&mut rt,json!({"op":"resumeSession","handle":1,"expectedRevision":1}))["error"]["code"],"staleRevision");
        assert_eq!(send(&mut rt,json!({"op":"destroy","handle":1,"expectedRevision":1}))["error"]["code"],"staleRevision");
        assert_eq!((rt.active_sessions(),rt.parked_sessions(),rt.parked_weight()),(0,1,weight));
        assert_eq!(send(&mut rt,json!({"op":"destroy","handle":1,"expectedRevision":0}))["status"],"ok");
        assert_eq!((rt.active_sessions(),rt.parked_sessions(),rt.parked_weight()),(0,0,0));
        assert_eq!(send(&mut rt,json!({"op":"create","documentId":"virtual-ime","source":"y"}))["result"]["handle"],2);
    }
    #[test]
    fn parking_releases_active_capacity_without_recycling_handles() {
        let mut rt=Runtime::default();
        for index in 0..MAX_SESSIONS {
            let handle=index as u32+1;
            assert_eq!(send(&mut rt,json!({"op":"create","documentId":format!("virtual-{index}"),"source":"x"}))["result"]["handle"],handle);
            assert_eq!(send(&mut rt,json!({"op":"parkSession","handle":handle,"expectedRevision":0}))["status"],"ok");
        }
        assert_eq!(rt.active_sessions(),0);assert_eq!(rt.parked_sessions(),MAX_SESSIONS);
        assert_eq!(send(&mut rt,json!({"op":"create","documentId":"visible","source":"z"}))["result"]["handle"],65);
        assert_eq!(rt.active_sessions(),1);assert_eq!(rt.parked_sessions(),MAX_SESSIONS);
    }
}
