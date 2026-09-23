//! Accent attachment and annotation stacking use the injected OpenType MATH
//! metrics. Real font outlines/assemblies are reused; no browser math layout.
use std::sync::Arc;
use crate::accents::{self,Mark};
use crate::error::{EngineError,EngineResult,ErrorCode};
use crate::model::Node;
use crate::scene::Rect;
use crate::session::NodeId;
use super::{MathStyle,boxes::{Builder,Fragment}};

impl Builder<'_> {
    pub fn accent(&mut self,id:NodeId,command:&str,body:&Node,style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment>{
        let spec=accents::spec(command).ok_or_else(||EngineError::new(ErrorCode::UnsupportedSyntax,"Unknown accent"))?;
        let scale=style.scale(self.math);
        let base=self.node(body,style,cramped||!spec.below,depth+1)?;
        let top=base.ink.map_or(-base.ascent,|r|r.y_min);
        let bottom=base.ink.map_or(base.descent,|r|r.y_max);
        let base_attachment=base.accent.unwrap_or(base.advance/2.0);
        let mut result=Fragment::new(id,"accent");result.italic=base.italic;
        match spec.mark {
            Mark::Rule=>{
                let (gap,thickness,extra)=if spec.below {(self.math.underbar_gap,self.math.underbar_rule,self.math.underbar_extra)}else{(self.math.overbar_gap,self.math.overbar_rule,self.math.overbar_extra)};
                let y=if spec.below {bottom+gap*scale}else{top-(gap+thickness)*scale};
                let width=base.advance;
                result.accent=base.accent;result.child(base,0.0,0.0);
                result.rule(Rect::new(0.0,y,width,y+thickness*scale));
                if spec.below {result.descent+=extra*scale;}else{result.ascent+=extra*scale;}
            }
            Mark::Segment=>{
                let (gap,rule,extra)=if spec.below {(self.math.underbar_gap,self.math.underbar_rule,self.math.underbar_extra)}
                    else{(self.math.overbar_gap,self.math.overbar_rule,self.math.overbar_extra)};
                let thickness=rule*scale;
                if ![gap,rule,extra].into_iter().all(|x|x.is_finite()&&x>=0.0)||thickness<=0.0 {
                    return Err(EngineError::new(ErrorCode::InvalidInput,"Invalid line segment mark metrics"));
                }
                // Authored geometric rule, not a substitute glyph or copied
                // renderer path: short perpendicular ticks bound a line segment.
                let cap_height=(4.0*thickness).max(0.18*scale);
                let logical=Rect::new(base.advance.min(0.0),top,base.advance.max(0.0),bottom);
                let extent=base.ink.map_or(logical,|ink|logical.union(ink));
                let left=extent.x_min;let right=extent.x_max.max(left+2.0*thickness);
                let center=if spec.below {bottom+gap*scale+cap_height/2.0}else{top-gap*scale-cap_height/2.0};
                result.accent=base.accent;result.child(base,0.0,0.0);
                result.rule(Rect::new(left,center-thickness/2.0,right,center+thickness/2.0));
                result.rule(Rect::new(left,center-cap_height/2.0,left+thickness,center+cap_height/2.0));
                result.rule(Rect::new(right-thickness,center-cap_height/2.0,right,center+cap_height/2.0));
                if spec.below{result.descent+=extra*scale;}else{result.ascent+=extra*scale;}
            }
            Mark::Glyph{codepoint,stretch:wide}=>{
                let (font,mark)=self.horizontal_product_or_math_glyph(
                    codepoint,
                    wide.then_some((base.advance/scale).max(1e-6)),
                )?;
                // Narrow accents follow the actual single-glyph attachment.
                // Extended decorations are centered over the whole base box.
                let attachment=if wide||spec.below {base.advance/2.0}else{base_attachment};
                let mark_anchor=if wide||spec.below {mark.width/2.0}else{mark.attachment};
                let x=attachment-mark_anchor*scale;
                let shift=(-x).max(0.0);
                let mark_y=if spec.below {
                    bottom+(if spec.ornament {self.math.underbar_gap*scale}else{0.0})-mark.ink_top*scale
                }else if spec.ornament {
                    top-self.math.overbar_gap*scale-mark.ink_bottom*scale
                }else{
                    -(-top).max(self.math.accent_base_height*scale)-mark.ink_bottom*scale
                };
                let width=(base.advance).max(x+mark.width*scale)+shift;
                result.accent=Some(base_attachment+shift);
                result.child(base,shift,0.0);
                for piece in mark.pieces {result.glyph(font,piece.shape,shift+x+piece.x*scale,mark_y,scale);}
                result.advance=width;
                if spec.below {result.descent+=self.math.underbar_extra*scale;}else{result.ascent+=self.math.overbar_extra*scale;}
            }
        }
        Ok(result)
    }

    pub fn annotation(&mut self,id:NodeId,_command:&str,above:bool,annotation:&Node,base:&Node,
        style:MathStyle,cramped:bool,depth:usize)->EngineResult<Fragment>{
        let base=self.node(base,style,cramped,depth+1)?;
        let label=self.node(annotation,style.script_child(),!above||cramped,depth+1)?;
        let (sup,sub)=if above{(Some(label),None)}else{(None,Some(label))};
        Ok(self.stack(id,"annotation",base,sup,sub,style,false))
    }
    pub(super) fn stack(&self,id:NodeId,role:&'static str,base:Arc<Fragment>,sup:Option<Arc<Fragment>>,
        sub:Option<Arc<Fragment>>,style:MathStyle,operator:bool)->Fragment {
        let scale=style.scale(self.math);
        let top=base.ink.map_or(base.ascent,|r|-r.y_min);
        let bottom=base.ink.map_or(base.descent,|r|r.y_max);
        let italic=if operator{base.italic/2.0}else{0.0};
        let sup_x=sup.as_ref().map_or(0.0,|s|(base.advance-s.advance)/2.0+italic);
        let sub_x=sub.as_ref().map_or(0.0,|s|(base.advance-s.advance)/2.0-italic);
        let min_x=0.0f64.min(sup_x).min(sub_x);let shift=-min_x;
        let mut f=Fragment::new(id,role);f.accent=base.accent.map(|a|a+shift);f.italic=base.italic;
        f.child(base,shift,0.0);
        if let Some(s)=sup {
            let gap=if operator{self.math.upper_limit_gap}else{self.math.overbar_gap}*scale;
            let rise=if operator{self.math.upper_limit_rise*scale}else{0.0};
            let y=-top-(gap+s.descent).max(rise);f.child(s,sup_x+shift,y);
            if !operator{f.ascent+=self.math.overbar_extra*scale;}
        }
        if let Some(s)=sub {
            let gap=if operator{self.math.lower_limit_gap}else{self.math.underbar_gap}*scale;
            let drop=if operator{self.math.lower_limit_drop*scale}else{0.0};
            let y=bottom+(gap+s.ascent).max(drop);f.child(s,sub_x+shift,y);
            if !operator{f.descent+=self.math.underbar_extra*scale;}
        }
        f
    }
}
