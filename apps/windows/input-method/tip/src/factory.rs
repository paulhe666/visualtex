//! The class factory COM asks for the text service.

use crate::service::TextService;
use std::ffi::c_void;
use windows::core::{implement, IUnknown, Interface, Ref, BOOL, GUID, HRESULT};
use windows::Win32::Foundation::CLASS_E_NOAGGREGATION;
use windows::Win32::System::Com::{IClassFactory, IClassFactory_Impl};

#[implement(IClassFactory)]
struct Factory;

impl IClassFactory_Impl for Factory_Impl {
    fn CreateInstance(
        &self,
        outer: Ref<'_, IUnknown>,
        interface: *const GUID,
        object: *mut *mut c_void,
    ) -> windows::core::Result<()> {
        if outer.is_some() {
            return Err(CLASS_E_NOAGGREGATION.into());
        }
        let service: IUnknown = TextService::new().into();
        unsafe { service.query(interface, object) }.ok()
    }

    fn LockServer(&self, _lock: BOOL) -> windows::core::Result<()> {
        Ok(())
    }
}

pub(crate) fn class_object(interface: *const GUID, object: *mut *mut c_void) -> HRESULT {
    let factory: IClassFactory = Factory.into();
    unsafe { factory.query(interface, object) }
}
