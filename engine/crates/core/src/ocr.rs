//! Deferred OCR is an insertion request, not a content/history owner. The
//! recognized payload remains LaTeX. Targets are strict, session-owned and
//! one-shot; no stale anchor is guessed or redirected to a current document.
use serde::{Deserialize,Serialize};
use crate::{EngineError,EngineResult,ErrorCode,positions::{Position,Selection,TreeIndex},session::{EngineSession,Revision}};
#[derive(Debug,Clone,PartialEq,Eq,Serialize,Deserialize)]
#[serde(rename_all="camelCase",deny_unknown_fields)]
pub struct OcrInsertionTarget {
    pub session_handle:u32,
    pub request_id:u32,
    pub document_id:String,
    pub revision:Revision,
    pub source_id:u32,
    pub selection_epoch:u32,
    pub selection:Selection,
}
pub(crate) struct OcrState {pub pending:Option<OcrInsertionTarget>,next_request_id:u32}
impl Default for OcrState{fn default()->Self{Self{pending:None,next_request_id:1}}}
fn conflict()->EngineError{EngineError::new(ErrorCode::Conflict,"OCR insertion target expired or belongs to another request; keep the recognized LaTeX for explicit insertion")}
impl EngineSession {
    pub fn capture_ocr_target(&mut self,handle:u32,expected:Revision)->EngineResult<OcrInsertionTarget>{
        self.check_revision(expected)?;self.ensure_no_composition()?;
        if handle==0{return Err(conflict());}
        if !self.document.complete{return Err(EngineError::new(ErrorCode::IncompleteSource,"Resolve the source draft before capturing an OCR target"));}
        if !self.document.supported{return Err(EngineError::new(ErrorCode::UnsupportedSyntax,"OCR cannot replace an unsupported visual draft"));}
        if matches!(&self.editor.selection,Selection::Range{anchor:Position::Text{..},..}|Selection::Range{focus:Position::Text{..},..}){return Err(EngineError::new(ErrorCode::InvalidPosition,"Choose a mathematical insertion point for OCR, not a text-run interior"));}
        TreeIndex::new(&self.document.tree)?.validate_selection(&self.editor.selection)?;
        let next=self.editor.ocr.next_request_id.checked_add(1).ok_or_else(||EngineError::new(ErrorCode::ResourceLimit,"OCR request identity exhausted"))?;
        let target=OcrInsertionTarget{session_handle:handle,request_id:self.editor.ocr.next_request_id,document_id:self.document_id().to_owned(),revision:self.revision,source_id:self.document.source.id(),selection_epoch:self.editor.selection_epoch,selection:self.editor.selection.clone()};
        self.editor.ocr.next_request_id=next;self.editor.ocr.pending=Some(target.clone());Ok(target)
    }
    pub fn cancel_ocr_target(&mut self,request_id:u32)->bool{
        if self.editor.ocr.pending.as_ref().is_some_and(|p|p.request_id==request_id){self.editor.ocr.pending=None;true}else{false}
    }
    pub(crate) fn invalidate_ocr_target(&mut self){self.editor.ocr.pending=None;}
    pub(crate) fn validate_ocr_target(&self,target:&OcrInsertionTarget)->EngineResult<()>{
        self.ensure_no_composition()?;self.check_revision(target.revision)?;
        if self.editor.ocr.pending.as_ref()!=Some(target)||target.document_id!=self.document_id()||target.source_id!=self.document.source.id()
            ||target.selection_epoch!=self.editor.selection_epoch||target.selection!=self.editor.selection{return Err(conflict());}
        TreeIndex::new(&self.document.tree)?.validate_selection(&target.selection)?;Ok(())
    }
}
