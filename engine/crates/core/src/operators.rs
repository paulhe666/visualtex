//! Authored n-ary operators. The ordered tail retains scripts and placement
//! directives exactly; effective placement is derived, never written over source.
use serde::{Deserialize,Serialize};
use crate::model::{Node,NodeKind,ScriptSlot};

pub const MAX_OPERATOR_PARTS:usize=128;
#[derive(Debug,Clone,Copy,PartialEq,Eq,Serialize,Deserialize)]
#[serde(rename_all="camelCase")]
pub enum LimitMode { Auto, Limits, NoLimits, DisplayLimits }
impl LimitMode {
    pub fn command(self)->Option<&'static str>{match self{
        Self::Auto=>None,Self::Limits=>Some("\\limits"),Self::NoLimits=>Some("\\nolimits"),Self::DisplayLimits=>Some("\\displaylimits"),
    }}
    pub fn from_command(command:&str)->Option<Self>{match command{
        "\\limits"=>Some(Self::Limits),"\\nolimits"=>Some(Self::NoLimits),"\\displaylimits"=>Some(Self::DisplayLimits),_=>None,
    }}
}
#[derive(Debug,Clone,Serialize)]
#[serde(tag="kind",rename_all="camelCase",rename_all_fields="camelCase")]
pub enum OperatorPart {
    Script {slot:ScriptSlot},
    Limits {before:String,command:String},
}
pub fn slots(parts:&[OperatorPart])->impl Iterator<Item=&ScriptSlot>+Clone{
    parts.iter().filter_map(|part|if let OperatorPart::Script{slot}=part{Some(slot)}else{None})
}
pub fn slots_mut(parts:&mut [OperatorPart])->impl Iterator<Item=&mut ScriptSlot>{
    parts.iter_mut().filter_map(|part|if let OperatorPart::Script{slot}=part{Some(slot)}else{None})
}
pub fn navigation_slots(parts:&[OperatorPart])->Vec<&ScriptSlot>{
    // Product default: lower -> upper -> following expression.
    crate::model::script_navigation_order(slots(parts))
}
/// Both fixed glyph operators and editable operator bodies use one ordered tail.
pub fn parts(node:&Node)->Option<&[OperatorPart]>{match &node.kind {
    NodeKind::LargeOperator{parts,..}|NodeKind::GeneralOperator{parts,..}=>Some(parts),
    NodeKind::CustomSymbol{asset,parts,..} if asset.role()==crate::model::AtomClass::Operator=>Some(parts),_=>None,
}}
pub fn parts_mut(node:&mut Node)->Option<&mut Vec<OperatorPart>>{match &mut node.kind {
    NodeKind::LargeOperator{parts,..}|NodeKind::GeneralOperator{parts,..}=>Some(parts),
    NodeKind::CustomSymbol{asset,parts,..} if asset.role()==crate::model::AtomClass::Operator=>Some(parts),_=>None,
}}
pub fn valid_parts(parts:&[OperatorPart])->bool{
    parts.len()<=MAX_OPERATOR_PARTS && parts.iter().all(|p|match p{
        OperatorPart::Limits{command,..}=>LimitMode::from_command(command).is_some(),
        OperatorPart::Script{slot}=>slot.argument.is_row(),
    })
}
pub fn effective_mode(parts:&[OperatorPart])->LimitMode{
    parts.iter().rev().find_map(|part|match part{
        OperatorPart::Limits{command,..}=>LimitMode::from_command(command),_=>None,
    }).unwrap_or(LimitMode::Auto)
}
#[derive(Debug,Clone,Copy,PartialEq,Eq)]
pub enum OperatorMark { Glyph(char), Text(&'static str) }
#[derive(Debug,Clone,Copy)]
pub struct OperatorSpec {pub mark:OperatorMark,pub display_limits:bool}
pub fn spec(command:&str)->Option<OperatorSpec>{
    use OperatorMark::*;
    let (mark,display_limits)=match command{
        "\\sum"|"∑"=>(Glyph('∑'),true),"\\prod"|"∏"=>(Glyph('∏'),true),"\\coprod"|"∐"=>(Glyph('∐'),true),
        "\\bigcup"|"⋃"=>(Glyph('⋃'),true),"\\bigcap"|"⋂"=>(Glyph('⋂'),true),
        "\\bigsqcup"|"⨆"=>(Glyph('⨆'),true),"\\biguplus"|"⨄"=>(Glyph('⨄'),true),
        "\\bigvee"|"⋁"=>(Glyph('⋁'),true),"\\bigwedge"|"⋀"=>(Glyph('⋀'),true),
        "\\bigodot"|"⨀"=>(Glyph('⨀'),true),"\\bigoplus"|"⨁"=>(Glyph('⨁'),true),"\\bigotimes"|"⨂"=>(Glyph('⨂'),true),
        "\\int"|"∫"=>(Glyph('∫'),false),"\\iint"|"∬"=>(Glyph('∬'),false),"\\iiint"|"∭"=>(Glyph('∭'),false),
        "\\oint"|"∮"=>(Glyph('∮'),false),"\\oiint"|"∯"=>(Glyph('∯'),false),"\\oiiint"|"∰"=>(Glyph('∰'),false),
        "\\lim"=>(Text("lim"),true),"\\limsup"=>(Text("lim sup"),true),"\\liminf"=>(Text("lim inf"),true),
        "\\min"=>(Text("min"),true),"\\max"=>(Text("max"),true),"\\inf"=>(Text("inf"),true),"\\sup"=>(Text("sup"),true),
        _=>(Glyph(crate::rare_operators::glyph(command)?),false),
    };Some(OperatorSpec{mark,display_limits})
}
pub fn stacked(spec:OperatorSpec,mode:LimitMode,display:bool)->bool{match mode{
    LimitMode::Auto=>spec.display_limits&&display,LimitMode::Limits=>true,LimitMode::NoLimits=>false,LimitMode::DisplayLimits=>display,
}}
