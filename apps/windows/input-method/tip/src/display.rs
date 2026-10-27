//! The look of pinyin being spelled in the field: underlined, as Windows' own input methods show
//! it (a display attribute the field draws). The formula's anchor has none.

use std::cell::Cell;
use windows::core::{implement, Result, BSTR, GUID};
use windows::Win32::Foundation::{E_INVALIDARG, S_FALSE};
use windows::Win32::UI::TextServices::{
    IEnumTfDisplayAttributeInfo, IEnumTfDisplayAttributeInfo_Impl, ITfDisplayAttributeInfo,
    ITfDisplayAttributeInfo_Impl, TF_ATTR_INPUT, TF_DISPLAYATTRIBUTE, TF_LS_SOLID,
};

/// The display attribute of text being composed.
pub(crate) const GUID_INPUT: GUID = GUID::from_u128(0x7c1e5b2a_93d4_4b61_8e0f_a2d6c3b8f514);

fn attribute() -> TF_DISPLAYATTRIBUTE {
    TF_DISPLAYATTRIBUTE { lsStyle: TF_LS_SOLID, bAttr: TF_ATTR_INPUT, ..Default::default() }
}

#[implement(ITfDisplayAttributeInfo)]
pub(crate) struct AttributeInfo;

impl ITfDisplayAttributeInfo_Impl for AttributeInfo_Impl {
    fn GetGUID(&self) -> Result<GUID> {
        Ok(GUID_INPUT)
    }

    fn GetDescription(&self) -> Result<BSTR> {
        Ok(BSTR::from("VisualTeX composition"))
    }

    fn GetAttributeInfo(&self, info: *mut TF_DISPLAYATTRIBUTE) -> Result<()> {
        if info.is_null() {
            return Err(E_INVALIDARG.into());
        }
        unsafe { *info = attribute() };
        Ok(())
    }

    fn SetAttributeInfo(&self, _info: *const TF_DISPLAYATTRIBUTE) -> Result<()> {
        Ok(())
    }

    fn Reset(&self) -> Result<()> {
        Ok(())
    }
}

#[implement(IEnumTfDisplayAttributeInfo)]
pub(crate) struct AttributeEnum {
    index: Cell<u32>,
}

impl AttributeEnum {
    pub(crate) fn new() -> Self {
        Self { index: Cell::new(0) }
    }
}

impl IEnumTfDisplayAttributeInfo_Impl for AttributeEnum_Impl {
    fn Clone(&self) -> Result<IEnumTfDisplayAttributeInfo> {
        Ok(AttributeEnum { index: Cell::new(self.index.get()) }.into())
    }

    fn Next(&self, count: u32, infos: *mut Option<ITfDisplayAttributeInfo>, fetched: *mut u32) -> Result<()> {
        let mut given = 0u32;
        if count > 0 && self.index.get() == 0 && !infos.is_null() {
            unsafe { *infos = Some(AttributeInfo.into()) };
            given = 1;
            self.index.set(1);
        }
        if !fetched.is_null() {
            unsafe { *fetched = given };
        }
        if given == count {
            Ok(())
        } else {
            Err(S_FALSE.into())
        }
    }

    fn Reset(&self) -> Result<()> {
        self.index.set(0);
        Ok(())
    }

    fn Skip(&self, count: u32) -> Result<()> {
        if count > 0 && self.index.get() == 0 {
            self.index.set(1);
            if count == 1 {
                return Ok(());
            }
        }
        Err(S_FALSE.into())
    }
}
