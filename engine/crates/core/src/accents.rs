//! Supported accent commands and their typographic roles. This table is shared
//! by the parser, structural commands and layout; unknown commands stay opaque.
#[derive(Debug,Clone,Copy,PartialEq,Eq)]
pub enum Mark { Glyph { codepoint:char, stretch:bool }, Rule, Segment }
#[derive(Debug,Clone,Copy,PartialEq,Eq)]
pub struct AccentSpec { pub below:bool, pub ornament:bool, pub mark:Mark }
pub fn spec(command:&str)->Option<AccentSpec> {
    let (below,ornament,mark)=match command {
        "\\hat"=>(false,false,Mark::Glyph{codepoint:'\u{302}',stretch:false}),
        "\\widehat"=>(false,false,Mark::Glyph{codepoint:'\u{302}',stretch:true}),
        "\\tilde"=>(false,false,Mark::Glyph{codepoint:'\u{303}',stretch:false}),
        "\\widetilde"=>(false,false,Mark::Glyph{codepoint:'\u{303}',stretch:true}),
        "\\vec"=>(false,false,Mark::Glyph{codepoint:'\u{20d7}',stretch:false}),
        "\\bar"=>(false,false,Mark::Glyph{codepoint:'\u{304}',stretch:false}),
        "\\dot"=>(false,false,Mark::Glyph{codepoint:'\u{307}',stretch:false}),
        "\\ddot"=>(false,false,Mark::Glyph{codepoint:'\u{308}',stretch:false}),
        "\\dddot"=>(false,false,Mark::Glyph{codepoint:'\u{20db}',stretch:false}),
        "\\ddddot"=>(false,false,Mark::Glyph{codepoint:'\u{20dc}',stretch:false}),
        "\\breve"=>(false,false,Mark::Glyph{codepoint:'\u{306}',stretch:false}),
        "\\check"=>(false,false,Mark::Glyph{codepoint:'\u{30c}',stretch:false}),
        "\\acute"=>(false,false,Mark::Glyph{codepoint:'\u{301}',stretch:false}),
        "\\grave"=>(false,false,Mark::Glyph{codepoint:'\u{300}',stretch:false}),
        "\\mathring"=>(false,false,Mark::Glyph{codepoint:'\u{30a}',stretch:false}),
        "\\overline"=>(false,true,Mark::Rule),"\\underline"=>(true,true,Mark::Rule),
        "\\overrightarrow"=>(false,false,Mark::Glyph{codepoint:'→',stretch:true}),
        "\\overleftarrow"=>(false,false,Mark::Glyph{codepoint:'←',stretch:true}),
        "\\overleftrightarrow"=>(false,false,Mark::Glyph{codepoint:'↔',stretch:true}),
        "\\underrightarrow"=>(true,false,Mark::Glyph{codepoint:'→',stretch:true}),
        "\\underleftarrow"=>(true,false,Mark::Glyph{codepoint:'←',stretch:true}),
        "\\underleftrightarrow"=>(true,false,Mark::Glyph{codepoint:'↔',stretch:true}),
        "\\overbrace"=>(false,true,Mark::Glyph{codepoint:'⏞',stretch:true}),
        "\\underbrace"=>(true,true,Mark::Glyph{codepoint:'⏟',stretch:true}),
        "\\overparen"|"\\overarc"=>(false,true,Mark::Glyph{codepoint:'⏜',stretch:true}),
        "\\underparen"|"\\underarc"=>(true,true,Mark::Glyph{codepoint:'⏝',stretch:true}),
        "\\overgroup"=>(false,true,Mark::Glyph{codepoint:'⏠',stretch:true}),
        "\\undergroup"=>(true,true,Mark::Glyph{codepoint:'⏡',stretch:true}),
        "\\overleftharpoon"=>(false,false,Mark::Glyph{codepoint:'↼',stretch:true}),
        "\\overrightharpoon"=>(false,false,Mark::Glyph{codepoint:'⇀',stretch:true}),
        "\\overlinesegment"=>(false,true,Mark::Segment),
        "\\underlinesegment"=>(true,true,Mark::Segment),
        _=>return None,
    };
    Some(AccentSpec{below,ornament,mark})
}
pub fn annotation_above(command:&str)->Option<bool> {
    match command {"\\overset"|"\\stackrel"|"\\stackbin"=>Some(true),"\\underset"=>Some(false),_=>None}
}
