//! Contained mathematical grids. Column identities are metadata, not caret
//! positions or source offsets. Rows/cells are independently identified nodes.
use serde::{Deserialize, Serialize};
use crate::{error::{EngineError,EngineResult,ErrorCode}, model::IdAllocator};

pub const MAX_GRID_ROWS: usize = 256;
pub const MAX_GRID_COLUMNS: usize = 64;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all="camelCase")]
pub enum GridColumnAlignment { Left, Center, Right }

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all="camelCase")]
pub struct ArrayColumnSpec {
    pub before_spec:String,
    pub source:String,
    pub alignments:Vec<GridColumnAlignment>,
    /// Boundary indices in 0..=column_count. Repeated entries represent
    /// repeated authored rules; the source spelling remains authoritative.
    pub vertical_rules:Vec<usize>,
}
impl ArrayColumnSpec {
    pub fn parse(before_spec:String,source:String)->EngineResult<Self>{
        if source.len()>512{return Err(EngineError::new(ErrorCode::ResourceLimit,"Array column specification exceeds its bounded source budget"));}
        let mut alignments=Vec::new();let mut vertical_rules=Vec::new();
        for ch in source.chars(){match ch{
            'l'=>alignments.push(GridColumnAlignment::Left),'c'=>alignments.push(GridColumnAlignment::Center),'r'=>alignments.push(GridColumnAlignment::Right),
            '|'=>vertical_rules.push(alignments.len()),
            ch if ch.is_ascii_whitespace()=>{},
            _=>return Err(EngineError::new(ErrorCode::UnsupportedSyntax,"Array column specification supports only l/c/r alignment and | rules")),
        }}
        if alignments.is_empty()||alignments.len()>MAX_GRID_COLUMNS{return Err(EngineError::new(ErrorCode::InvalidInput,"Array column specification must define a bounded nonempty column set"));}
        Ok(Self{before_spec,source,alignments,vertical_rules})
    }
    pub fn validate(&self)->EngineResult<()> {
        let parsed=Self::parse(self.before_spec.clone(),self.source.clone())?;
        if parsed.alignments!=self.alignments||parsed.vertical_rules!=self.vertical_rules{return Err(EngineError::new(ErrorCode::InvalidInput,"Array column metadata does not match its authored specification"));}
        Ok(())
    }
    pub fn canonical_source(&self)->String {
        let mut out=String::new();
        for boundary in 0..=self.alignments.len(){
            for _ in self.vertical_rules.iter().filter(|&&at|at==boundary){out.push('|');}
            if let Some(alignment)=self.alignments.get(boundary){out.push(match alignment{GridColumnAlignment::Left=>'l',GridColumnAlignment::Center=>'c',GridColumnAlignment::Right=>'r'});}
        }
        out
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(transparent)]
pub struct GridColumnId(pub u32);
impl GridColumnId {
    pub fn allocate(ids: &mut IdAllocator) -> EngineResult<Self> { Ok(Self(ids.allocate()?.0)) }
}

/// Each supported environment owns its column and layout policy. Environment
/// spellings remain explicit so structural edits preserve authored begin/end.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum GridKind {
    Gathered, Matrix, Cases, Aligned, Align, AlignStar, Array,
    #[serde(rename="pmatrix")] PMatrix,
    #[serde(rename="bmatrix")] BMatrix,
    #[serde(rename="Bmatrix")] BraceMatrix,
    #[serde(rename="vmatrix")] VMatrix,
    #[serde(rename="Vmatrix")] DoubleVMatrix,
    /// amsmath \substack{a\\b}: a command, not an environment, whose rows
    /// are centered in the ambient (usually script) style.
    Substack,
}
impl GridKind {
    pub fn name(self) -> &'static str { match self {
        Self::Gathered=>"gathered", Self::Matrix=>"matrix", Self::Cases=>"cases",
        Self::Aligned=>"aligned", Self::Align=>"align", Self::AlignStar=>"align*", Self::Array=>"array",
        Self::PMatrix=>"pmatrix", Self::BMatrix=>"bmatrix", Self::BraceMatrix=>"Bmatrix",
        Self::VMatrix=>"vmatrix", Self::DoubleVMatrix=>"Vmatrix", Self::Substack=>"substack",
    } }
    pub fn is_aligned(self)->bool { matches!(self,Self::Aligned|Self::Align|Self::AlignStar) }
    pub fn is_matrix(self)->bool {matches!(self,Self::Matrix|Self::PMatrix|Self::BMatrix|Self::BraceMatrix|Self::VMatrix|Self::DoubleVMatrix)}
    pub fn is_array(self)->bool {self==Self::Array}
    pub fn matrix_delimiters(self)->Option<(&'static str,&'static str)>{match self{
        Self::PMatrix=>Some(("(",")")), Self::BMatrix=>Some(("[","]")), Self::BraceMatrix=>Some(("\\lbrace","\\rbrace")),
        Self::VMatrix=>Some(("\\lvert","\\rvert")), Self::DoubleVMatrix=>Some(("\\lVert","\\rVert")), _=>None,
    }}
    pub fn from_name(name: &str) -> Option<Self> {
        match name { "gathered"=>Some(Self::Gathered), "matrix"=>Some(Self::Matrix), "cases"=>Some(Self::Cases),
            "aligned"=>Some(Self::Aligned), "align"=>Some(Self::Align), "align*"=>Some(Self::AlignStar), "array"=>Some(Self::Array),
            "pmatrix"=>Some(Self::PMatrix), "bmatrix"=>Some(Self::BMatrix), "Bmatrix"=>Some(Self::BraceMatrix),
            "vmatrix"=>Some(Self::VMatrix), "Vmatrix"=>Some(Self::DoubleVMatrix), _=>None }
    }
    pub fn accepts_columns(self, count: usize) -> bool {
        if self==Self::Gathered||self==Self::Substack {return count==1;} if self==Self::Cases {return (1..=2).contains(&count);}
        (self.is_matrix()||self.is_aligned()||self.is_array())&&(1..=MAX_GRID_COLUMNS).contains(&count)
    }
}
