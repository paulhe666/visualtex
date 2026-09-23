//! Pure input-method composition state. Real platform event ordering is a P06
//! acceptance gate; these types alone do not demonstrate system IME support.
use serde::Serialize;
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::positions::Selection;
use crate::session::Revision;
use crate::source::MAX_SOURCE_BYTES;
use crate::transactions::InputMode;

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Composition {
    pub id: u32,
    pub revision: Revision,
    pub selection_epoch: u32,
    pub selection: Selection,
    pub mode: InputMode,
    pub preedit: String,
}
impl Composition {
    pub fn validate_target(&self, revision: Revision, selection_epoch: u32, selection: &Selection) -> EngineResult<()> {
        if self.revision != revision {
            return Err(EngineError::new(ErrorCode::StaleRevision, "Composition belongs to an older document revision"));
        }
        if self.selection_epoch != selection_epoch || self.selection != *selection {
            return Err(EngineError::new(ErrorCode::Conflict, "Composition target selection changed"));
        }
        Ok(())
    }
    pub fn validate_text(text: &str) -> EngineResult<()> {
        if text.len() > MAX_SOURCE_BYTES {
            Err(EngineError::new(ErrorCode::ResourceLimit, "Composition preedit exceeds the source budget"))
        } else { Ok(()) }
    }
}
