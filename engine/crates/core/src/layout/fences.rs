//! Symmetric stretchable fences around a mathematical body, centered on MATH axis.
use crate::{delimiters,stretch};
use crate::error::{EngineError,EngineResult,ErrorCode};
use crate::model::{AtomClass,Node};
use crate::session::NodeId;
use super::{MathStyle,boxes::{Builder,Fragment}};

impl Builder<'_> {
    pub fn fence(&mut self,id:NodeId,left:&str,body:&Node,right:&str,
        style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment> {
        let previous_middle_target=self.middle_target;
        self.middle_target=None;
        let measured=match self.node(body,style,cramped,depth+1){Ok(body)=>body,Err(error)=>{self.middle_target=previous_middle_target;return Err(error);}};
        let scale=style.scale(self.math);let axis=self.math.axis*scale;
        let target=2.0*(measured.ascent-axis).max(measured.descent+axis).max(1e-6);
        self.middle_target=Some(target);
        let body_result=self.node(body,style,cramped,depth+1);
        self.middle_target=previous_middle_target;
        let body=body_result?;
        let shape=|source:&str,class:AtomClass|->EngineResult<Option<(&crate::fonts::FontAsset,stretch::StretchedGlyph)>> {
            let glyph=delimiters::glyph(source).ok_or_else(||EngineError::new(ErrorCode::UnsupportedSyntax,"Unsupported fence delimiter"))?;
            glyph.map(|ch|self.vertical_product_or_math_glyph(ch,class,target,scale)).transpose()
        };
        let left=shape(left,AtomClass::Open)?;let right=shape(right,AtomClass::Close)?;
        let left_width=left.as_ref().map_or(0.0,|(_,s)|s.width*scale);
        let right_x=left_width+body.advance;
        let mut result=Fragment::new(id,"fence");
        result.italic=body.italic;
        result.child(body,left_width,0.0);
        for (glyph,x) in [(left,0.0),(right,right_x)] {
            if let Some((font,glyph))=glyph {
                let top=-axis-glyph.height*scale/2.0;
                result.advance=result.advance.max(x+glyph.width*scale);
                result.italic=glyph.italic_correction*scale;
                for piece in glyph.pieces { result.glyph(font,piece.shape,x,top+piece.y*scale,scale); }
            }
        }
        Ok(result)
    }
}
