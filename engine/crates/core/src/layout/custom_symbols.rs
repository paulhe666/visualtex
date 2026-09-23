//! Asset metrics and math roles participate in the same box/limit machinery as
//! built-in symbols. Artwork is never put into the font glyph resource table.
use std::sync::Arc;
use crate::{EngineError,EngineResult,ErrorCode,custom_symbols::{SymbolAsset,SymbolLimits},model::AtomClass,operators::{LimitMode,OperatorPart},session::NodeId};
use super::{MathStyle,boxes::{Builder,Fragment}};
impl Builder<'_>{
    pub(super) fn custom_symbol(&mut self,id:NodeId,asset:&SymbolAsset,parts:&[OperatorPart],style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment>{
        asset.ensure_supported()?;
        let m=asset.metrics();let scale=style.scale(self.math);let mut f=Fragment::new(id,"customSymbol");
        f.advance=m.width_em*scale;f.ascent=m.ascent_em*scale;f.descent=m.descent_em*scale;
        for (i,shape) in asset.geometry().iter().enumerate(){let resolved=if let Some(text)=&shape.text_source{
            let binding=self.custom_text_fonts.iter().find(|b|b.family==text.font_family&&b.style==text.font_style&&b.weight==text.font_weight)
                .ok_or_else(||EngineError::new(ErrorCode::InvalidInput,"Custom text artwork requires an exact explicit registered font binding"))?;
            Arc::new(crate::custom_symbols::materialize_text(shape,binding.font,m.ascent_em)?)
        }else{Arc::clone(shape)};f.artwork(asset,i,resolved,scale);}
        if asset.role()!=AtomClass::Operator{return Ok(f);}
        let stacked=match crate::operators::effective_mode(parts){
            LimitMode::Limits=>true,LimitMode::NoLimits=>false,LimitMode::DisplayLimits=>style==MathStyle::Display,
            LimitMode::Auto=>match asset.limits(){SymbolLimits::Auto=>style==MathStyle::Display,SymbolLimits::Limits=>true,SymbolLimits::Nolimits=>false},
        };
        self.operator_limits(f,parts,style,cramped,depth,stacked)
    }
}
