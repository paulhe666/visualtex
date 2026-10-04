//! Layout for the bounded compatibility-macro subset. Source macros remain
//! MacroCall nodes; these routines interpret their mathematical presentation
//! from the same injected fonts/MATH metrics as native structures.
use std::sync::Arc;
use crate::error::{EngineError,EngineResult,ErrorCode};
use crate::macros::BuiltinMacroKind;
use crate::model::{AtomClass,Node};
use crate::scene::Rect;
use crate::session::NodeId;
use crate::styles::MathVariant;
use super::{MathStyle,boxes::{Builder,Fragment}};

impl Builder<'_> {
    pub(super) fn builtin_macro(&mut self,id:NodeId,kind:BuiltinMacroKind,args:&[Arc<Node>],
        style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment>{
        use BuiltinMacroKind::*;
        match kind {
            MathRelation|MathBinary=>{
                let child=self.node(&args[0],style,cramped,depth+1)?;
                let mut f=Fragment::new(id,"macroCall");
                f.italic=child.italic;f.accent=child.accent;f.child(child,0.0,0.0);Ok(f)
            },
            Absolute=>self.fence(id,"|",&args[0],"|",style,cramped,depth),
            Norm=>self.fence(id,"\\lVert",&args[0],"\\rVert",style,cramped,depth),
            Bra=>self.fence(id,"\\langle",&args[0],"\\rvert",style,cramped,depth),
            Ket=>self.fence(id,"\\lvert",&args[0],"\\rangle",style,cramped,depth),
            Expectation=>self.fence(id,"\\langle",&args[0],"\\rangle",style,cramped,depth),
            Differential=>self.prefixed_argument(id,'d',&args[0],style,cramped,depth),
            Derivative=>self.derivative_macro(id,'d',&args[0],&args[1],style,cramped,depth),
            PartialDerivative=>self.derivative_macro(id,'∂',&args[0],&args[1],style,cramped,depth),
            Commutator=>self.delimited_macro(id,"[",
                &[(None,&args[0]),(Some(','),&args[1])],"]",style,cramped,depth),
            Anticommutator|PoissonBracket=>self.delimited_macro(id,"\\lbrace",
                &[(None,&args[0]),(Some(','),&args[1])],"\\rbrace",style,cramped,depth),
            BraKet=>self.delimited_macro(id,"\\langle",
                &[(None,&args[0]),(Some('|'),&args[1])],"\\rangle",style,cramped,depth),
            MatrixElement=>self.delimited_macro(id,"\\langle",
                &[(None,&args[0]),(Some('|'),&args[1]),(Some('|'),&args[2])],"\\rangle",style,cramped,depth),
            KetBra=>self.ket_bra(id,&args[0],&args[1],style,cramped,depth),
            BoldVector=>self.variant_macro(id,MathVariant::Bold,&args[0],style,cramped,depth),
            ArrowVector=>self.accented_variant_macro(id,MathVariant::Bold,"\\vec",&args[0],style,cramped,depth),
            UnitVector=>self.accented_variant_macro(id,MathVariant::Bold,"\\hat",&args[0],style,cramped,depth),
        }
    }

    fn variant_macro(&mut self,id:NodeId,variant:MathVariant,arg:&Node,style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment>{
        let previous=self.variant;self.variant=variant;let child=self.node(arg,style,cramped,depth+1);self.variant=previous;
        let child=child?;let mut f=Fragment::new(id,"macroCall");f.italic=child.italic;f.accent=child.accent;f.child(child,0.0,0.0);Ok(f)
    }
    fn accented_variant_macro(&mut self,id:NodeId,variant:MathVariant,command:&str,arg:&Node,style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment>{
        let previous=self.variant;self.variant=variant;let result=self.accent(id,command,arg,style,cramped,depth+1);self.variant=previous;result
    }
    fn prefixed_argument(&mut self,id:NodeId,prefix:char,arg:&Node,style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment>{
        let scale=style.scale(self.math);let child=self.node(arg,style,cramped,depth+1)?;
        let (font,gid)=self.math_glyph(prefix)?;let shape=font.glyph(gid)?;let prefix_width=shape.advance*scale;
        let mut f=Fragment::new(id,"macroCall");f.glyph(font,shape,0.0,0.0,scale);f.child(child,prefix_width,0.0);f.advance=f.advance.max(prefix_width);Ok(f)
    }
    fn derivative_macro(&mut self,id:NodeId,prefix:char,numerator:&Node,denominator:&Node,
        style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment>{
        let scale=style.scale(self.math);let child_style=style.fraction_child();let child_scale=child_style.scale(self.math);
        let num=self.node(numerator,child_style,cramped,depth+1)?;let den=self.node(denominator,child_style,true,depth+1)?;
        let (font,gid)=self.math_glyph(prefix)?;let shape=font.glyph(gid)?;let prefix_width=shape.advance*child_scale;
        let thin=if prefix=='∂'{0.08*child_scale}else{0.04*child_scale};
        let nw=prefix_width+thin+num.advance;let dw=prefix_width+thin+den.advance;
        let m=self.math;let axis=m.axis*scale;let thickness=m.fraction_rule*scale;let display=style==MathStyle::Display;
        let num_gap=(if display{m.numerator_display_gap}else{m.numerator_gap})*scale;
        let den_gap=(if display{m.denominator_display_gap}else{m.denominator_gap})*scale;
        let u=((if display{m.numerator_display_shift}else{m.numerator_shift})*scale).max(num.descent+axis+thickness/2.0+num_gap);
        let v=((if display{m.denominator_display_shift}else{m.denominator_shift})*scale).max(den.ascent-axis+thickness/2.0+den_gap);
        let width=nw.max(dw)+0.2*scale;let nx=(width-nw)/2.0;let dx=(width-dw)/2.0;
        let mut f=Fragment::new(id,"macroCall");
        f.glyph(font,Arc::clone(&shape),nx,-u,child_scale);f.child(num,nx+prefix_width+thin,-u);
        f.glyph(font,shape,dx,v,child_scale);f.child(den,dx+prefix_width+thin,v);
        f.advance=width;f.rule(Rect::new(0.0,-axis-thickness/2.0,width,-axis+thickness/2.0));Ok(f)
    }
    fn delimiter_glyph(&self,source:&str,target:f64,scale:f64)->EngineResult<(crate::stretch::StretchedGlyph,&crate::fonts::FontAsset)> {
        let ch=crate::delimiters::glyph(source).ok_or_else(||EngineError::new(ErrorCode::UnsupportedSyntax,"Unsupported macro delimiter"))?
            .ok_or_else(||EngineError::new(ErrorCode::UnsupportedSyntax,"Hidden macro delimiter is not valid here"))?;
        let (font,stretched)=self.vertical_product_or_math_glyph(ch,AtomClass::Inner,target.max(0.01),scale)?;
        Ok((stretched,font))
    }
    fn delimited_macro(&mut self,id:NodeId,left:&str,parts:&[(Option<char>,&Arc<Node>)],right:&str,
        style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment>{
        let scale=style.scale(self.math);let mut rendered=Vec::with_capacity(parts.len());let mut ascent=0.0_f64;let mut descent=0.0_f64;
        for (_,node) in parts {let child=self.node(node,style,cramped,depth+1)?;ascent=ascent.max(child.ascent);descent=descent.max(child.descent);rendered.push(child);}
        let axis=self.math.axis*scale;let target=2.0*(ascent-axis).max(descent+axis).max(scale);
        let (lg,lf)=self.delimiter_glyph(left,target,scale)?;let (rg,rf)=self.delimiter_glyph(right,target,scale)?;
        let mut f=Fragment::new(id,"macroCall");let top=-axis-lg.height*scale/2.0;let mut x=0.0;
        for piece in lg.pieces {f.glyph(lf,piece.shape,x,top+piece.y*scale,scale);}x+=lg.width*scale;
        for ((separator,_),child) in parts.iter().zip(rendered.into_iter()) {
            if let Some(ch)=separator {
                let (font,gid)=self.math_glyph(*ch)?;let shape=font.glyph(gid)?;f.glyph(font,Arc::clone(&shape),x,0.0,scale);x+=shape.advance*scale+0.08*scale;
            }
            let advance=child.advance;f.child(child,x,0.0);x+=advance;
        }
        let right_top=-axis-rg.height*scale/2.0;for piece in rg.pieces {f.glyph(rf,piece.shape,x,right_top+piece.y*scale,scale);}x+=rg.width*scale;
        f.advance=x;Ok(f)
    }
    fn ket_bra(&mut self,id:NodeId,ket:&Node,bra:&Node,style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment>{
        let scale=style.scale(self.math);let a=self.node(ket,style,cramped,depth+1)?;let b=self.node(bra,style,cramped,depth+1)?;
        let ascent=a.ascent.max(b.ascent);let descent=a.descent.max(b.descent);let axis=self.math.axis*scale;
        let target=2.0*(ascent-axis).max(descent+axis).max(scale);let mut f=Fragment::new(id,"macroCall");let mut x=0.0;
        for (source,child,closing) in [("\\lvert",a,"\\rangle"),("\\langle",b,"\\rvert")] {
            let (lg,lf)=self.delimiter_glyph(source,target,scale)?;let top=-axis-lg.height*scale/2.0;for p in lg.pieces{f.glyph(lf,p.shape,x,top+p.y*scale,scale);}x+=lg.width*scale;
            let advance=child.advance;f.child(child,x,0.0);x+=advance;
            let (rg,rf)=self.delimiter_glyph(closing,target,scale)?;let top=-axis-rg.height*scale/2.0;for p in rg.pieces{f.glyph(rf,p.shape,x,top+p.y*scale,scale);}x+=rg.width*scale;
            if closing=="\\rangle" {x=(x-scale/6.0).max(0.0);}
        }
        f.advance=x;Ok(f)
    }
}
