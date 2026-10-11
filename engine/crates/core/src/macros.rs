//! Bounded built-in compatibility macros. These are explicit language data,
//! not TeX expansion and never execute user code. Authored command spelling and
//! argument nodes remain in the model; layout interprets only these fixed kinds.
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum BuiltinMacroKind {
    Absolute,
    Norm,
    Differential,
    Derivative,
    PartialDerivative,
    Commutator,
    Anticommutator,
    PoissonBracket,
    Bra,
    Ket,
    BraKet,
    Expectation,
    MatrixElement,
    KetBra,
    BoldVector,
    ArrowVector,
    UnitVector,
    MathRelation,
    MathBinary,
    /// amsmath \xrightarrow{label} (mandatory label form).
    ExtensibleRightArrow,
    /// amsmath \xleftarrow{label} (mandatory label form).
    ExtensibleLeftArrow,
    /// amsmath \pmod{n}: "(mod n)" with leading space.
    ParenthesizedMod,
}

#[derive(Debug, Clone, Copy)]
pub struct MacroSpec {
    pub command: &'static str,
    pub kind: BuiltinMacroKind,
    pub arity: usize,
}

const SPECS: &[MacroSpec] = &[
    MacroSpec { command: "\\mathrel", kind: BuiltinMacroKind::MathRelation, arity: 1 },
    MacroSpec { command: "\\mathbin", kind: BuiltinMacroKind::MathBinary, arity: 1 },
    MacroSpec { command: "\\abs", kind: BuiltinMacroKind::Absolute, arity: 1 },
    MacroSpec { command: "\\norm", kind: BuiltinMacroKind::Norm, arity: 1 },
    MacroSpec { command: "\\dd", kind: BuiltinMacroKind::Differential, arity: 1 },
    MacroSpec { command: "\\dv", kind: BuiltinMacroKind::Derivative, arity: 2 },
    MacroSpec { command: "\\pdv", kind: BuiltinMacroKind::PartialDerivative, arity: 2 },
    MacroSpec { command: "\\comm", kind: BuiltinMacroKind::Commutator, arity: 2 },
    MacroSpec { command: "\\acomm", kind: BuiltinMacroKind::Anticommutator, arity: 2 },
    MacroSpec { command: "\\pb", kind: BuiltinMacroKind::PoissonBracket, arity: 2 },
    MacroSpec { command: "\\bra", kind: BuiltinMacroKind::Bra, arity: 1 },
    MacroSpec { command: "\\ket", kind: BuiltinMacroKind::Ket, arity: 1 },
    MacroSpec { command: "\\braket", kind: BuiltinMacroKind::BraKet, arity: 2 },
    MacroSpec { command: "\\expval", kind: BuiltinMacroKind::Expectation, arity: 1 },
    MacroSpec { command: "\\mel", kind: BuiltinMacroKind::MatrixElement, arity: 3 },
    MacroSpec { command: "\\ketbra", kind: BuiltinMacroKind::KetBra, arity: 2 },
    MacroSpec { command: "\\vb", kind: BuiltinMacroKind::BoldVector, arity: 1 },
    MacroSpec { command: "\\va", kind: BuiltinMacroKind::ArrowVector, arity: 1 },
    MacroSpec { command: "\\vu", kind: BuiltinMacroKind::UnitVector, arity: 1 },
    MacroSpec { command: "\\xrightarrow", kind: BuiltinMacroKind::ExtensibleRightArrow, arity: 1 },
    MacroSpec { command: "\\xleftarrow", kind: BuiltinMacroKind::ExtensibleLeftArrow, arity: 1 },
    MacroSpec { command: "\\pmod", kind: BuiltinMacroKind::ParenthesizedMod, arity: 1 },
];

pub fn spec(command: &str) -> Option<MacroSpec> {
    SPECS.iter().copied().find(|spec| spec.command == command)
}

pub fn all() -> &'static [MacroSpec] { SPECS }
