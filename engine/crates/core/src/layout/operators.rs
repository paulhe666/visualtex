//! N-ary glyph selection and side/stacked bounds share font metrics with the
//! existing script engine. The operator owns its glyphs; bounds are child rows.
use std::sync::Arc;
use crate::{error::{EngineError,EngineResult,ErrorCode},model::{AtomClass,ScriptKind},
    operators::{self,OperatorMark,OperatorPart},session::NodeId};
use super::{MathStyle,boxes::{Builder,Fragment}};

impl Builder<'_> {
    pub fn large_operator(&mut self,id:NodeId,command:&str,parts:&[OperatorPart],style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment>{
        let spec=operators::spec(command).ok_or_else(||EngineError::new(ErrorCode::UnsupportedSyntax,"Unknown large operator"))?;
        let scale=style.scale(self.math);let mut f=Fragment::new(id,"largeOperator");
        match spec.mark {
            OperatorMark::Glyph(ch)=>{
                // TeX operators are designed Size1/Size2 glyphs, not a scaled
                // ordinary letter. Keep text faces away from the MATH-only API.
                let tex=self.tex_glyph(if style==MathStyle::Display {"size2"} else {"size1"},ch)?;
                let preferred=if style==MathStyle::Display {None} else {self.product_symbol_glyph(ch,AtomClass::Operator)?};
                let (font,gid)=match tex.or(preferred) {Some(value)=>value,None=>self.math_glyph(ch)?};
                let glyph=if tex.is_some() || font.math.is_none() {font.glyph(gid)?} else {font.operator_glyph(gid,style==MathStyle::Display)?};
                let ink=glyph.ink.ok_or_else(||EngineError::new(ErrorCode::InvalidInput,"Operator glyph has no ink"))?;
                let y=(-(ink[1]+ink[3])/2.0-self.math.axis)*scale;
                f.advance=glyph.advance*scale;f.italic=glyph.italic_correction*scale;
                f.glyph(font,glyph,0.0,y,scale);
            }
            OperatorMark::Text(text)=>{
                let font=self.alphabet_upright_font.unwrap_or(self.math_font);
                let run=font.shape_text(text)?;f.advance=run.advance*scale;
                for g in run.glyphs{f.glyph(font,g.shape,g.x*scale,g.y*scale,scale);}
            }
        }
        let stacked=operators::stacked(spec,operators::effective_mode(parts),style==MathStyle::Display);
        self.operator_limits(f,parts,style,cramped,depth,stacked)
    }

    pub fn general_operator(&mut self,id:NodeId,body:&crate::model::Node,parts:&[OperatorPart],
        style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment>{
        let child=self.node(body,style,cramped,depth+1)?;
        let (advance,italic,accent)=(child.advance,child.italic,child.accent);
        let mut f=Fragment::new(id,"generalOperator");
        f.child(child,0.0,0.0);f.advance=advance;f.italic=italic;f.accent=accent;
        // An authored mathop body is a genuine mathematical box, not the
        // following operand and not an implicitly enlarged replacement glyph.
        let stacked=match operators::effective_mode(parts){
            operators::LimitMode::Auto|operators::LimitMode::DisplayLimits=>style==MathStyle::Display,
            operators::LimitMode::Limits=>true,operators::LimitMode::NoLimits=>false,
        };
        self.operator_limits(f,parts,style,cramped,depth,stacked)
    }

    pub(super) fn operator_limits(&mut self,mut f:Fragment,parts:&[OperatorPart],style:MathStyle,
        cramped:bool,depth:usize,stacked:bool)->EngineResult<Fragment>{
        let scale=style.scale(self.math);
        let mut sup:Option<Arc<Fragment>>=None;let mut sub:Option<Arc<Fragment>>=None;
        for slot in operators::slots(parts){
            let b=self.node(&slot.argument,style.script_child(),cramped||slot.script==ScriptKind::Subscript,depth+1)?;
            if slot.script==ScriptKind::Superscript{sup=Some(b);}else{sub=Some(b);}
        }
        if sup.is_none()&&sub.is_none(){return Ok(f);}
        if stacked{
            let (width,ascent,descent,italic)=(f.advance,f.ascent,f.descent,f.italic);
            let sx=sup.as_ref().map_or(0.0,|s|(width-s.advance)/2.0+italic/2.0);
            let dx=sub.as_ref().map_or(0.0,|s|(width-s.advance)/2.0-italic/2.0);
            let shift=-sx.min(dx).min(0.0);f.translate_x(shift);
            // Lower-first scene order follows the mathematical edit navigation;
            // y positions, not sequence or source offsets, determine display.
            if let Some(b)=sub{
                let y=descent+(self.math.lower_limit_gap*scale+b.ascent).max(self.math.lower_limit_drop*scale);
                f.child(b,dx+shift,y);
            }
            if let Some(b)=sup{
                let y=ascent+(self.math.upper_limit_gap*scale+b.descent).max(self.math.upper_limit_rise*scale);
                f.child(b,sx+shift,-y);
            }
        }else{
            let (u,v)=self.script_shifts(&f,sup.as_deref(),sub.as_deref(),style,cramped,false);
            let (x,italic)=(f.advance,f.italic);
            if let Some(b)=sub{f.child(b,x,v);}
            if let Some(b)=sup{f.child(b,x+italic,-u);}
            f.advance+=self.math.space_after_script*scale;
        }
        // Following ordinary atoms are spaced as operators by their parent row;
        // an attached upper bound's italic correction is not an operator margin.
        f.italic=0.0;Ok(f)
    }
}
