//! Contained gathered layout: shared column width and explicit row baselines.
//! No DOM table measurements or separate per-line editor instances are used.
use std::sync::Arc;
use crate::{error::{EngineError,EngineResult,ErrorCode},grids::{ArrayColumnSpec,GridColumnAlignment,GridColumnId,GridKind},
    model::{AtomClass,Node,NodeKind},scene::{GridColumnAnchor,Rect},session::NodeId};
use super::{MathStyle,boxes::{Builder,Fragment}};

fn extent(fragment:&Fragment)->Rect {
    let logical=Rect::new(fragment.advance.min(0.0),-fragment.ascent,fragment.advance.max(0.0),fragment.descent);
    fragment.ink.map_or(logical,|ink|logical.union(ink))
}
impl Builder<'_>{
    pub(super) fn grid_row(&mut self,id:NodeId,cells:&[Arc<Node>],style:MathStyle,
        cramped:bool,depth:usize)->EngineResult<Fragment>{
        if cells.len()!=1 {return Err(EngineError::new(ErrorCode::UnsupportedSyntax,"This gathered slice requires one cell per row"));}
        let child=self.node(&cells[0],style,cramped,depth+1)?;
        let advance=child.advance;let mut row=Fragment::new(id,"gridRow");row.child(child,0.0,0.0);row.advance=advance;Ok(row)
    }
    pub(super) fn grid(&mut self,id:NodeId,environment:GridKind,columns:&[GridColumnId],array_spec:Option<&ArrayColumnSpec>,rows:&[Arc<Node>],
        ambient:MathStyle,depth:usize)->EngineResult<Fragment>{
        if !environment.accepts_columns(columns.len()) || rows.is_empty(){
            return Err(EngineError::new(ErrorCode::InvalidInput,"Invalid gathered dimensions"));
        }
        if environment!=GridKind::Gathered {return self.cell_grid(id,environment,columns,array_spec,rows,ambient,depth);}
        // gathered is a display-style subenvironment. The parent mathematical
        // axis still determines its vertical placement as one inline box.
        let row_style=MathStyle::Display;
        let local=row_style.scale(self.math);
        // Documented grid defaults, not claims of MATH constants: 0.8/0.2em
        // minimum row strut and 0.3em clear inter-row gap. Taller content wins.
        let gap=0.3*local;let mut fragments=Vec::with_capacity(rows.len());let mut width:f64=0.0;
        let mut height=0.0;let mut previous_descent=0.0;
        for (i,row) in rows.iter().enumerate(){
            let fragment=self.node(row,row_style,false,depth+1)?;let bounds=extent(&fragment);
            width=width.max(bounds.width());
            let ascent=fragment.ascent.max(0.8*local);let descent=fragment.descent.max(0.2*local);
            let baseline=if i==0{ascent}else{height+previous_descent+gap+ascent};
            height=baseline;previous_descent=descent;
            fragments.push((fragment,bounds,baseline));
        }
        height+=previous_descent;
        let axis=self.math.axis*ambient.scale(self.math);let shift=-height/2.0-axis;
        let mut result=Fragment::new(id,"grid");
        for (row,bounds,baseline) in fragments {
            let x=(width-bounds.width())/2.0-bounds.x_min;
            result.child(row,x,baseline+shift);
        }
        result.advance=width;result.ascent=height/2.0+axis;result.descent=(height/2.0-axis).max(0.0);
        result.column_anchors.push(GridColumnAnchor{grid_id:id,column_id:columns[0],index:0,x:0.0,width});
        Ok(result)
    }

    fn cell_grid(&mut self,id:NodeId,environment:GridKind,columns:&[GridColumnId],array_spec:Option<&ArrayColumnSpec>,rows:&[Arc<Node>],
        ambient:MathStyle,depth:usize)->EngineResult<Fragment>{
        // Widths and alignment belong to the environment, not each separate
        // row. Aligned pairs alternate right/left in display math; Matrix/Cases
        // use text math. Contextual right-cell glue is never written to source.
        let aligned=environment.is_aligned();
        let array_spec=if environment==GridKind::Array {
            let spec=array_spec.ok_or_else(||EngineError::new(ErrorCode::InvalidInput,"Array layout requires an explicit column specification"))?;
            spec.validate()?;if spec.alignments.len()!=columns.len(){return Err(EngineError::new(ErrorCode::InvalidInput,"Array layout column count does not match its specification"));}Some(spec)
        }else{if array_spec.is_some(){return Err(EngineError::new(ErrorCode::InvalidInput,"Non-array grid cannot carry array layout metadata"));}None};
        let style=if aligned{MathStyle::Display}else{MathStyle::Text};let scale=style.scale(self.math);
        let mut widths=vec![0.0_f64;columns.len()];let mut measured=Vec::with_capacity(rows.len());
        for row in rows {
            let NodeKind::GridRow{cells,..}=&row.kind else {
                return Err(EngineError::new(ErrorCode::InvalidInput,"Matrix child is not a grid row"));
            };
            if cells.len()!=columns.len(){return Err(EngineError::new(ErrorCode::InvalidInput,"Matrix cells do not match its columns"));}
            let mut content=Vec::with_capacity(cells.len());
            let mut ascent=0.8*scale;let mut descent=0.2*scale;
            for (column,cell) in cells.iter().enumerate(){
                let fragment=if aligned&&column%2==1 {
                    self.built+=1;Arc::new(self.math_row(cell,style,false,depth+2,Some(AtomClass::Ordinary))?)
                }else{self.node(cell,style,false,depth+2)?};let bounds=extent(&fragment);
                widths[column]=widths[column].max(bounds.width());
                ascent=ascent.max(-bounds.y_min);descent=descent.max(bounds.y_max);
                content.push((fragment,bounds));
            }
            measured.push((row.id,content,ascent,descent));
        }
        let row_gap=0.3*scale;
        let mut starts=Vec::with_capacity(columns.len());let mut width=0.0;
        for (i,cell_width) in widths.iter().enumerate(){
            starts.push(width);width+=cell_width;
            if i+1<columns.len(){width+=if aligned{if i%2==1{2.0*scale}else{0.0}}else{scale};}
        }
        let height: f64=measured.iter().map(|(_,_,a,d)|a+d).sum::<f64>()+row_gap*(rows.len()-1)as f64;
        let axis=self.math.axis*ambient.scale(self.math);let mut y=-height/2.0-axis;
        let mut result=Fragment::new(id,"grid");
        for (row_id,cells,ascent,descent) in measured {
            y+=ascent;let mut row=Fragment::new(row_id,"gridRow");self.built+=1;
            for (column,(cell,bounds)) in cells.into_iter().enumerate(){
                let alignment=if aligned{if column%2==0{widths[column]-bounds.width()}else{0.0}}
                    else if environment==GridKind::Cases{0.0}
                    else if let Some(spec)=array_spec {match spec.alignments[column]{GridColumnAlignment::Left=>0.0,GridColumnAlignment::Center=>(widths[column]-bounds.width())/2.0,GridColumnAlignment::Right=>widths[column]-bounds.width()}}
                    else{(widths[column]-bounds.width())/2.0};
                let x=starts[column]+alignment-bounds.x_min;
                row.child(cell,x,0.0);
            }
            row.advance=width;row.ascent=ascent;row.descent=descent;
            result.child(Arc::new(row),0.0,y);y+=descent+row_gap;
        }
        result.advance=width;result.ascent=height/2.0+axis;result.descent=(height/2.0-axis).max(0.0);
        for (index,column_id) in columns.iter().enumerate(){
            result.column_anchors.push(GridColumnAnchor{grid_id:id,column_id:*column_id,index,x:starts[index],width:widths[index]});
        }
        if let Some(spec)=array_spec {
            let thickness=(self.math.fraction_rule*scale).max(0.02*scale);
            let mut grouped=std::collections::BTreeMap::<usize,usize>::new();for &boundary in &spec.vertical_rules{*grouped.entry(boundary).or_default()+=1;}
            for (boundary,count) in grouped {
                let center=if boundary==0{0.0}else if boundary==columns.len(){width}else{(starts[boundary-1]+widths[boundary-1]+starts[boundary])/2.0};
                for index in 0..count {let offset=(index as f64-(count as f64-1.0)/2.0)*2.2*thickness;result.rule(Rect::new(center+offset-thickness/2.0,-result.ascent,center+offset+thickness/2.0,result.descent));}
            }
        }
        if environment==GridKind::Cases {
            // The brace is a genuine font construction owned by the same grid.
            // It is not a CSS border or a separate editable fence/host.
            let brace_scale=ambient.scale(self.math);
            let (brace_font,glyph)=self.vertical_product_or_math_glyph('{',crate::model::AtomClass::Open,height,brace_scale)?;
            result.translate_x(glyph.width*brace_scale+0.15*brace_scale);
            let top=-axis-glyph.height*brace_scale/2.0;
            for piece in glyph.pieces {result.glyph(brace_font,piece.shape,0.0,top+piece.y*brace_scale,brace_scale);}
        } else if let Some((left,right))=environment.matrix_delimiters() {
            // Matrix wrappers are presentation owned by this same Grid. They
            // never create a Fence node/editor. Moving the content fragment also
            // moves all column anchors and caret geometry by the real glyph width.
            let fence_scale=ambient.scale(self.math);let target=height.max(fence_scale);
            let left_ch=crate::delimiters::glyph(left).and_then(|g|g).ok_or_else(||EngineError::new(ErrorCode::UnsupportedSyntax,"Invalid matrix left delimiter"))?;
            let right_ch=crate::delimiters::glyph(right).and_then(|g|g).ok_or_else(||EngineError::new(ErrorCode::UnsupportedSyntax,"Invalid matrix right delimiter"))?;
            let (left_font,left_glyph)=self.vertical_product_or_math_glyph(left_ch,crate::model::AtomClass::Open,target,fence_scale)?;
            let (right_font,right_glyph)=self.vertical_product_or_math_glyph(right_ch,crate::model::AtomClass::Close,target,fence_scale)?;
            let gap=0.15*fence_scale;let left_width=left_glyph.width*fence_scale;
            result.translate_x(left_width+gap);
            let left_top=-axis-left_glyph.height*fence_scale/2.0;
            for piece in left_glyph.pieces {result.glyph(left_font,piece.shape,0.0,left_top+piece.y*fence_scale,fence_scale);}
            let right_x=result.advance+gap;let right_top=-axis-right_glyph.height*fence_scale/2.0;
            for piece in right_glyph.pieces {result.glyph(right_font,piece.shape,right_x,right_top+piece.y*fence_scale,fence_scale);}
            result.advance=right_x+right_glyph.width*fence_scale;
        }
        Ok(result)
    }
}
