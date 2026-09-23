use std::sync::Arc;
use serde::{Deserialize, Serialize};
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::model::{self, IdAllocator, LatexMode, Node};
use crate::parser::{self, ParseResult};
use crate::source::SourceText;
use crate::source_edit::{self, SourceChangeMap, SourceEdit};

/// Local to a document; never an Office formulaId, source offset or JS pointer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(transparent)]
pub struct NodeId(pub u32);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(transparent)]
pub struct Revision(pub u32);

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SessionSnapshot {
    pub document_id: String,
    pub revision: Revision,
    pub root_id: NodeId,
    pub source: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentSnapshot {
    pub document_id: String,
    pub revision: Revision,
    pub source_id: u32,
    pub source: String,
    pub analysis: Arc<ParseResult>,
    /// Only present for an incomplete draft with a previously complete tree.
    /// Its older source identity is explicit; it is not the current source parse.
    pub retained_tree: Option<Arc<Node>>,
    /// Never offered as a replacement for incomplete/unsupported source.
    pub canonical_latex: Option<String>,
}

/// The only writer of a document. Parsing and later transactions extend this
/// owner rather than creating a second frontend source/history state store.
pub struct EngineSession {
    document_id: String,
    pub(crate) revision: Revision,
    pub(crate) root_id: NodeId,
    pub(crate) next_node_id: u32,
    pub(crate) next_source_id: u32,
    pub(crate) document: Arc<ParseResult>,
    pub(crate) last_complete: Option<Arc<Node>>,
    pub(crate) editor: crate::editing::EditorState,
}

impl EngineSession {
    pub fn new(document_id: String) -> EngineResult<Self> {
        if document_id.trim().is_empty() || document_id.len() > 256 || document_id.chars().any(char::is_control) {
            return Err(EngineError::new(ErrorCode::InvalidInput, "documentId must be a nonempty opaque identifier of at most 256 UTF-8 bytes"));
        }
        let root_id = NodeId(1);
        let mut ids = IdAllocator::new(2);
        let document = Arc::new(parser::parse(SourceText::new(1, "")?, root_id, &mut ids)?);
        Ok(Self { document_id, revision: Revision(0), root_id, next_node_id: ids.next_unused(),
            next_source_id: 2, document, last_complete: None,
            editor: crate::editing::EditorState::new(root_id, 0) })
    }

    pub fn document_id(&self) -> &str { &self.document_id }
    pub fn revision(&self) -> Revision { self.revision }
    pub fn source(&self) -> &str { self.document.source.text() }

    pub fn check_revision(&self, expected: Revision) -> EngineResult<()> {
        if self.revision != expected {
            return Err(EngineError::new(ErrorCode::StaleRevision, "Request revision does not match the document"));
        }
        Ok(())
    }

    pub fn allocate_node_id(&mut self) -> EngineResult<NodeId> {
        let mut ids = IdAllocator::new(self.next_node_id);
        let id = ids.allocate()?;
        self.next_node_id = ids.next_unused();
        Ok(id)
    }

    pub fn snapshot(&self) -> SessionSnapshot {
        SessionSnapshot { document_id: self.document_id.clone(), revision: self.revision,
            root_id: self.root_id, source: self.source().to_owned() }
    }

    pub fn document_snapshot(&self) -> DocumentSnapshot {
        DocumentSnapshot {
            document_id: self.document_id.clone(), revision: self.revision,
            source_id: self.document.source.id(), source: self.source().to_owned(),
            analysis: Arc::clone(&self.document),
            retained_tree: if self.document.complete { None } else { self.last_complete.clone() },
            canonical_latex: if self.document.complete && self.document.supported {
                Some(model::latex(&self.document.tree, LatexMode::Canonical))
            } else { None },
        }
    }

    pub fn patch_source(&mut self, expected: Revision, edit: SourceEdit) -> EngineResult<(DocumentSnapshot, SourceChangeMap)> {
        self.check_revision(expected)?;
        self.ensure_no_composition()?;
        let mut ids = IdAllocator::new(self.next_node_id);
        let prepared = source_edit::prepare(&self.document, &edit, self.next_source_id, &mut ids)?;
        if !prepared.changed { return Ok((self.document_snapshot(), prepared.mapping)); }
        self.commit_document(prepared.document, ids.next_unused(), None, None, None, false, false)?;
        Ok((self.document_snapshot(), prepared.mapping))
    }

    /// Explicit full-source replacement, used for opening/importing source.
    /// Incremental source/editor transactions will use this same owner's commit
    /// boundary; full replacement does not pretend all inner node IDs survived.
    pub fn replace_source(&mut self, expected: Revision, source: String) -> EngineResult<DocumentSnapshot> {
        self.check_revision(expected)?;
        self.ensure_no_composition()?;
        if source == self.source() { return Ok(self.document_snapshot()); }
        let source = SourceText::new(self.next_source_id, source)?;
        let mut ids = IdAllocator::new(self.next_node_id);
        let document = Arc::new(parser::parse_with_symbols(source, self.root_id, &mut ids, Arc::clone(&self.document.symbol_assets))?);
        // Source and visual editing share the same atomic history/selection owner.
        // Invalid syntax remains a draft; resource failures do not commit it.
        self.commit_document(document, ids.next_unused(), None, None, None, false, false)?;
        Ok(self.document_snapshot())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn identities_are_monotonic_and_not_text_hashes() {
        let mut s = EngineSession::new("doc-a".into()).unwrap();
        assert_eq!(s.allocate_node_id().unwrap(), NodeId(2));
        assert_eq!(s.allocate_node_id().unwrap(), NodeId(3));
        assert_eq!(s.snapshot().root_id, NodeId(1));
    }
    #[test]
    fn validates_document_identity_without_normalizing_it() {
        for id in ["", "   ", "a\0b"] { assert!(EngineSession::new(id.into()).is_err()); }
        assert!(EngineSession::new("x".repeat(257)).is_err());
        assert_eq!(EngineSession::new("文档-A".into()).unwrap().document_id(), "文档-A");
    }
    #[test]
    fn source_replacement_has_one_writer_and_noop_does_not_dirty() {
        let mut s = EngineSession::new("doc".into()).unwrap();
        let first = s.replace_source(Revision(0), "x_1".into()).unwrap();
        assert_eq!(first.revision, Revision(1));
        let first_json = serde_json::to_value(&first).unwrap();
        assert_eq!(serde_json::to_value(s.replace_source(Revision(1), "x_1".into()).unwrap()).unwrap(), first_json);
        assert_eq!(s.replace_source(Revision(0), "y".into()).unwrap_err().code, ErrorCode::StaleRevision);
        assert_eq!(s.snapshot().source, "x_1");
    }
    #[test]
    fn incomplete_source_is_retained_with_explicit_previous_tree() {
        let mut s = EngineSession::new("doc".into()).unwrap();
        let complete = s.replace_source(Revision(0), "x+1".into()).unwrap();
        let draft = s.replace_source(Revision(1), "\\frac{x}{".into()).unwrap();
        assert!(!draft.analysis.complete);
        assert_eq!(draft.source, "\\frac{x}{");
        assert!(draft.canonical_latex.is_none());
        assert_eq!(model::latex(draft.retained_tree.as_ref().unwrap(), LatexMode::Lossless), complete.source);
        assert_eq!(model::latex(&draft.analysis.tree, LatexMode::Lossless), draft.source);
        let fixed = s.replace_source(Revision(2), "\\frac{x}{2}".into()).unwrap();
        assert!(fixed.analysis.exportable);
        assert!(fixed.retained_tree.is_none());
    }
    #[test]
    fn resource_failure_leaves_entire_session_unchanged() {
        let mut s = EngineSession::new("doc".into()).unwrap();
        s.replace_source(Revision(0), "x".into()).unwrap();
        let before = serde_json::to_value(s.document_snapshot()).unwrap();
        let next_node = s.next_node_id;
        let next_source = s.next_source_id;
        assert_eq!(s.replace_source(Revision(1), "x".repeat(9000)).unwrap_err().code, ErrorCode::ResourceLimit);
        assert_eq!(serde_json::to_value(s.document_snapshot()).unwrap(), before);
        assert_eq!(s.next_node_id, next_node);
        assert_eq!(s.next_source_id, next_source);
    }
}
