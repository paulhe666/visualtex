//! The look of pinyin being spelled in the field: underlined, as Windows' own input methods show
//! it (a display attribute the field draws). The formula's anchor gets one only in applications
//! typing through IMM32: it marks the anchor converted and selected, as the clause an input method
//! is converting, so that the field hides its caret there (Qt does) and draws nothing for it. The
//! space a field holds for a formula (space.rs) is input without a line: the field draws no
//! composition underline under the formula (VisualTeX covers what a field draws anyway).

use std::cell::Cell;
use windows::core::{implement, Result, BSTR, GUID};
use windows::Win32::Foundation::{E_INVALIDARG, S_FALSE};
use windows::Win32::UI::TextServices::{
    IEnumTfDisplayAttributeInfo, IEnumTfDisplayAttributeInfo_Impl, ITfDisplayAttributeInfo,
    ITfDisplayAttributeInfo_Impl, TF_ATTR_INPUT, TF_ATTR_TARGET_CONVERTED, TF_DISPLAYATTRIBUTE, TF_LS_NONE,
    TF_LS_SOLID,
};

/// The display attribute of text being composed.
pub(crate) const GUID_INPUT: GUID = GUID::from_u128(0x7c1e5b2a_93d4_4b61_8e0f_a2d6c3b8f514);
/// The display attribute of the formula's anchor (IMM32 applications).
pub(crate) const GUID_ANCHOR: GUID = GUID::from_u128(0x2f6b8d41_5ac7_4e93_b1d2_8c4e7a9f3b06);
/// The display attribute of the space a field holds for a formula.
pub(crate) const GUID_SPACE: GUID = GUID::from_u128(0x5d3c9e17_2b84_4f6a_9c01_7e2a4b8d6f35);

const ALL: [GUID; 3] = [GUID_INPUT, GUID_ANCHOR, GUID_SPACE];

fn attribute(guid: &GUID) -> TF_DISPLAYATTRIBUTE {
    if *guid == GUID_ANCHOR {
        // no colour and no line: nothing to draw for it
        TF_DISPLAYATTRIBUTE { lsStyle: TF_LS_NONE, bAttr: TF_ATTR_TARGET_CONVERTED, ..Default::default() }
    } else if *guid == GUID_SPACE {
        TF_DISPLAYATTRIBUTE { lsStyle: TF_LS_NONE, bAttr: TF_ATTR_INPUT, ..Default::default() }
    } else {
        TF_DISPLAYATTRIBUTE { lsStyle: TF_LS_SOLID, bAttr: TF_ATTR_INPUT, ..Default::default() }
    }
}

#[implement(ITfDisplayAttributeInfo)]
pub(crate) struct AttributeInfo {
    guid: GUID,
}

impl AttributeInfo {
    /// The attribute of `guid`, when it is one of this input method's.
    pub(crate) fn of(guid: &GUID) -> Option<Self> {
        ALL.contains(guid).then_some(Self { guid: *guid })
    }
}

impl ITfDisplayAttributeInfo_Impl for AttributeInfo_Impl {
    fn GetGUID(&self) -> Result<GUID> {
        Ok(self.guid)
    }

    fn GetDescription(&self) -> Result<BSTR> {
        Ok(BSTR::from(if self.guid == GUID_ANCHOR {
            "VisualTeX formula"
        } else if self.guid == GUID_SPACE {
            "VisualTeX formula space"
        } else {
            "VisualTeX composition"
        }))
    }

    fn GetAttributeInfo(&self, info: *mut TF_DISPLAYATTRIBUTE) -> Result<()> {
        if info.is_null() {
            return Err(E_INVALIDARG.into());
        }
        unsafe { *info = attribute(&self.guid) };
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
    index: Cell<usize>,
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
        if !infos.is_null() {
            while given < count && self.index.get() < ALL.len() {
                let info: ITfDisplayAttributeInfo = AttributeInfo { guid: ALL[self.index.get()] }.into();
                // an out array: write without dropping what the caller left in it
                unsafe { infos.add(given as usize).write(Some(info)) };
                given += 1;
                self.index.set(self.index.get() + 1);
            }
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
        let skipped = (count as usize).min(ALL.len() - self.index.get());
        self.index.set(self.index.get() + skipped);
        if skipped == count as usize {
            Ok(())
        } else {
            Err(S_FALSE.into())
        }
    }
}
