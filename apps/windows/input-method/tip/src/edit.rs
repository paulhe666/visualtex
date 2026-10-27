//! TSF edit sessions: document reads and writes run inside one, with the
//! edit cookie TSF hands to DoEditSession.

use windows::core::{implement, Result};
use windows::Win32::UI::TextServices::{
    ITfContext, ITfEditSession, ITfEditSession_Impl, TF_ES_ASYNCDONTCARE, TF_ES_READWRITE, TF_ES_SYNC,
};

#[implement(ITfEditSession)]
struct EditSession(Box<dyn Fn(u32) -> Result<()>>);

impl ITfEditSession_Impl for EditSession_Impl {
    fn DoEditSession(&self, cookie: u32) -> Result<()> {
        (self.0)(cookie)
    }
}

/// Runs `work` in a read-write edit session of `context` (now when TSF
/// allows it, else as soon as it can).
pub(crate) fn request(context: &ITfContext, client: u32, work: impl Fn(u32) -> Result<()> + 'static) -> Result<()> {
    let session: ITfEditSession = EditSession(Box::new(work)).into();
    let result = unsafe { context.RequestEditSession(client, &session, TF_ES_ASYNCDONTCARE | TF_ES_READWRITE)? };
    result.ok()
}

/// Runs `work` now (inside a key event TSF grants it), else as soon as it can.
pub(crate) fn request_now(context: &ITfContext, client: u32, work: impl Fn(u32) -> Result<()> + 'static) -> Result<()> {
    let session: ITfEditSession = EditSession(Box::new(work)).into();
    let now = unsafe { context.RequestEditSession(client, &session, TF_ES_SYNC | TF_ES_READWRITE) };
    match now {
        Ok(result) if result.is_ok() => Ok(()),
        _ => {
            let later = unsafe { context.RequestEditSession(client, &session, TF_ES_ASYNCDONTCARE | TF_ES_READWRITE)? };
            later.ok()
        }
    }
}
