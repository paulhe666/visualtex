//! Custom vector resources are not fake font glyphs. They retain their source
//! asset digest, explicit baseline-relative paths and ordered paint operations.
use std::{collections::BTreeSet,sync::Arc};
use serde::Serialize;
use crate::{EngineError,EngineResult,ErrorCode,custom_symbols::{ArtworkOperation,ArtworkPath},fonts::PathOp,paint::RgbColor,scene::{LayoutSnapshot,Rect},session::NodeId};
#[derive(Debug,Clone,Serialize)]
#[serde(rename_all="camelCase")]
pub struct ArtworkInstance {
    pub node_id:NodeId,pub asset_sha256:String,pub path_index:usize,
    pub shape:Arc<ArtworkPath>,pub x:f64,pub y:f64,pub scale:f64,
    #[serde(skip_serializing_if="Option::is_none")]pub color:Option<RgbColor>,
}
#[derive(Debug,Clone,Serialize)]
#[serde(tag="kind",rename_all="camelCase")]
pub enum PaintItem {Glyph{index:usize},Rule{index:usize},Artwork{index:usize}}
fn invalid()->EngineError{EngineError::new(ErrorCode::InvalidInput,"Invalid custom artwork scene resource or paint order")}
fn contained(a:Rect,b:Rect)->bool{b.x_min>=a.x_min-1e-8&&b.y_min>=a.y_min-1e-8&&b.x_max<=a.x_max+1e-8&&b.y_max<=a.y_max+1e-8}
fn transformed([a,b,c,d,e,f]:[f64;6],[x,y]:[f64;2])->Option<[f64;2]>{let p=[a*x+c*y+e,b*x+d*y+f];p.iter().all(|v|v.is_finite()&&v.abs()<=10_000.0).then_some(p)}
fn valid_matrix(m:[f64;6])->bool{m.iter().all(|v|v.is_finite()&&v.abs()<=1e8)&&(m[0]*m[3]-m[1]*m[2]).abs()>=1e-12}
fn transformed_rect(m:[f64;6],r:Rect)->Option<Rect>{let points=[transformed(m,[r.x_min,r.y_min])?,transformed(m,[r.x_max,r.y_min])?,transformed(m,[r.x_max,r.y_max])?,transformed(m,[r.x_min,r.y_max])?];Some(Rect::new(points.iter().map(|p|p[0]).fold(f64::INFINITY,f64::min),points.iter().map(|p|p[1]).fold(f64::INFINITY,f64::min),points.iter().map(|p|p[0]).fold(f64::NEG_INFINITY,f64::max),points.iter().map(|p|p[1]).fold(f64::NEG_INFINITY,f64::max)))}
fn same_f64(a:f64,b:f64)->bool{a.to_bits()==b.to_bits()}
fn validate_paint_order(scene:&LayoutSnapshot)->EngineResult<()> {
    let mut used=BTreeSet::new();let mut closed:BTreeSet<(NodeId,String)>=BTreeSet::new();
    let mut current:Option<(NodeId,String,usize,f64,f64,f64,Option<RgbColor>)>=None;
    for item in &scene.paint_order {
        let(key,index,limit)=match item{PaintItem::Glyph{index}=>(0,*index,scene.glyphs.len()),PaintItem::Rule{index}=>(1,*index,scene.rules.len()),PaintItem::Artwork{index}=>(2,*index,scene.artworks.len())};
        if index>=limit||!used.insert((key,index)){return Err(invalid());}
        let PaintItem::Artwork{index}=item else {if let Some((node,sha,..))=current.take(){closed.insert((node,sha));}continue};let a=&scene.artworks[*index];
        if let Some((node,sha,next,x,y,scale,color))=&mut current {
            if *node==a.node_id&&*sha==a.asset_sha256 {
                if a.path_index!=*next||!same_f64(*x,a.x)||!same_f64(*y,a.y)||!same_f64(*scale,a.scale)||*color!=a.color{return Err(invalid());}*next+=1;continue;
            }
            closed.insert((*node,sha.clone()));
        }
        let group=(a.node_id,a.asset_sha256.clone());if closed.contains(&group)||a.path_index!=0{return Err(invalid());}
        current=Some((a.node_id,a.asset_sha256.clone(),1,a.x,a.y,a.scale,a.color));
    }Ok(())
}
pub fn validate(scene:&LayoutSnapshot)->EngineResult<()> {
    if scene.artworks.is_empty(){if !scene.paint_order.is_empty(){return Err(invalid());}return Ok(());}
    if scene.artworks.len()>4096||scene.paint_order.len()!=scene.glyphs.len()+scene.rules.len()+scene.artworks.len(){return Err(invalid());}validate_paint_order(scene)?;
    let frame=Rect::new(0.,0.,scene.width_em,scene.height_em);let mut total=0usize;
    for artwork in &scene.artworks{
        if artwork.asset_sha256.len()!=64||!artwork.asset_sha256.bytes().all(|c|c.is_ascii_hexdigit()&&!c.is_ascii_uppercase())
            ||artwork.path_index>=4096*2||artwork.shape.path.is_empty()||artwork.shape.text_source.is_some()||!artwork.shape.ink.valid()||![artwork.x,artwork.y,artwork.scale,artwork.shape.stroke_width].iter().all(|v|v.is_finite())||artwork.scale<=0.0||!(0.0..=100.0).contains(&artwork.shape.stroke_width)
            ||artwork.shape.outline_font_id.as_ref().is_some_and(|id|!scene.fonts.iter().any(|font|font.id==*id))
            ||!valid_matrix(artwork.shape.transform)||!valid_matrix(artwork.shape.inner_transform)
            ||artwork.shape.clip.is_some_and(|r|!r.valid()||[r.x_min,r.y_min,r.x_max,r.y_max].iter().any(|v|v.abs()>10_000.0)){return Err(invalid());}
        let owner=scene.regions.iter().find(|r|r.node_id==artwork.node_id).ok_or_else(invalid)?;if owner.role!="customSymbol"{return Err(invalid());}
        total+=artwork.shape.path.len();if total>131072{return Err(EngineError::new(ErrorCode::ResourceLimit,"Scene custom artwork exceeds the path budget"));}
        let ink=artwork.shape.ink;let world=Rect::new(ink.x_min*artwork.scale+artwork.x,ink.y_min*artwork.scale+artwork.y,ink.x_max*artwork.scale+artwork.x,ink.y_max*artwork.scale+artwork.y);
        if artwork.shape.operation==ArtworkOperation::Paint&&(!contained(frame,world)||owner.ink.is_none_or(|r|!contained(r,world))){return Err(invalid());}
        if let Some(clip)=artwork.shape.clip{let clipped=transformed_rect(artwork.shape.transform,clip).ok_or_else(invalid)?;if !clipped.valid(){return Err(invalid());}}
        let mut has_move=false;for op in &artwork.shape.path{let p:&[f64]=match op{PathOp::Move(p)=>{has_move=true;p},PathOp::Line(p)=>p,PathOp::Quad(p)=>p,PathOp::Cubic(p)=>p,PathOp::Close=>&[]};if !has_move{return Err(invalid());}
            for xy in p.chunks_exact(2){if !xy.iter().all(|n|n.is_finite()&&n.abs()<=10_000.0){return Err(invalid());}let inner=transformed(artwork.shape.inner_transform,[xy[0],xy[1]]).ok_or_else(invalid)?;let p=transformed(artwork.shape.transform,inner).ok_or_else(invalid)?;if artwork.shape.clip.is_none()&&!contained(ink,Rect::new(p[0],p[1],p[0],p[1])){return Err(invalid());}}
        }
    }Ok(())
}
fn number(value:f64)->String{let mut v=format!("{value:.6}");while v.contains('.')&&v.ends_with('0'){v.pop();}if v.ends_with('.'){v.pop();}if v=="-0"{v="0".into();}v}
fn outline(out:&mut String,path:&[PathOp],x:f64,y:f64,scale:f64,color:Option<RgbColor>){
    if path.is_empty(){return;}out.push_str("<path d=\"");for op in path{
        let mut point=|prefix:&str,p:&[f64]|{out.push_str(prefix);for(i,v)in p.iter().enumerate(){if i>0{out.push(' ');}out.push_str(&number(v*scale+if i%2==0{x}else{y}));}};
        match op{PathOp::Move(p)=>point("M",p),PathOp::Line(p)=>point("L",p),PathOp::Quad(p)=>point("Q",p),PathOp::Cubic(p)=>point("C",p),PathOp::Close=>out.push('Z')}
    }out.push('"');if let Some(color)=color{out.push_str(&format!(" fill=\"{}\"",color.hex()));}out.push_str("/>");
}
fn matrix_text([a,b,c,d,e,f]:[f64;6])->String{format!("matrix({} {} {} {} {} {})",number(a),number(b),number(c),number(d),number(e),number(f))}
fn artwork_shape(out:&mut String,item:&ArtworkInstance,index:usize,paint:&str,scope:&str){
    if item.shape.path.is_empty(){return;}let outer=matrix_text(item.shape.transform);let inner=matrix_text(item.shape.inner_transform);out.push_str(&format!("<g transform=\"{}\">",outer));
    if let Some(clip)=item.shape.clip{let id=format!("vt-art-{scope}-clip-{index}");out.push_str(&format!("<defs><clipPath id=\"{}\" clipPathUnits=\"userSpaceOnUse\"><rect x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\"/></clipPath></defs><g clip-path=\"url(#{})\">",id,number(clip.x_min),number(clip.y_min),number(clip.width()),number(clip.height()),id));}
    out.push_str("<path d=\"");for op in &item.shape.path{let mut point=|prefix:&str,p:&[f64]|{out.push_str(prefix);for(i,v)in p.iter().enumerate(){if i>0{out.push(' ');}out.push_str(&number(*v));}};match op{PathOp::Move(p)=>point("M",p),PathOp::Line(p)=>point("L",p),PathOp::Quad(p)=>point("Q",p),PathOp::Cubic(p)=>point("C",p),PathOp::Close=>out.push('Z')}}
    out.push_str(&format!("\" transform=\"{}\"",inner));if item.shape.fill{out.push_str(&format!(" fill=\"{}\"",paint));}else{out.push_str(" fill=\"none\"");}
    if item.shape.stroke_width>0.0{out.push_str(&format!(" stroke=\"{}\" stroke-width=\"{}\" stroke-linecap=\"{}\" stroke-linejoin=\"{}\"",paint,number(item.shape.stroke_width),match item.shape.line_cap{crate::custom_symbols::ArtworkLineCap::Butt=>"butt",crate::custom_symbols::ArtworkLineCap::Round=>"round",crate::custom_symbols::ArtworkLineCap::Square=>"square"},match item.shape.line_join{crate::custom_symbols::ArtworkLineJoin::Miter=>"miter",crate::custom_symbols::ArtworkLineJoin::Round=>"round",crate::custom_symbols::ArtworkLineJoin::Bevel=>"bevel"}));}else{out.push_str(" stroke=\"none\"");}
    out.push_str("/>");if item.shape.clip.is_some(){out.push_str("</g>");}out.push_str("</g>");
}
fn artwork_group(out:&mut String,scene:&LayoutSnapshot,indices:&[usize],group_index:usize){
    let Some(first)=indices.first().map(|i|&scene.artworks[*i])else{return};let mut visible=None;let mut has_erase=false;
    for index in indices{let item=&scene.artworks[*index];if item.shape.operation==ArtworkOperation::Paint{visible=Some(visible.map_or(item.shape.ink,|r:Rect|r.union(item.shape.ink)));}else{has_erase=true;}}
    let Some(visible)=visible else{return};if visible.width()<=0.||visible.height()<=0.{return;}out.push_str(&format!("<g transform=\"translate({} {}) scale({})\">",number(first.x),number(first.y),number(first.scale)));
    if has_erase {let mask_id=format!("vt-art-mask-{group_index}");out.push_str(&format!("<defs><mask id=\"{}\" maskUnits=\"userSpaceOnUse\" maskContentUnits=\"userSpaceOnUse\" mask-type=\"luminance\" x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\">",mask_id,number(visible.x_min),number(visible.y_min),number(visible.width()),number(visible.height())));
        for index in indices{let item=&scene.artworks[*index];artwork_shape(out,item,*index,if item.shape.operation==ArtworkOperation::Paint{"#ffffff"}else{"#000000"},"mask");}out.push_str("</mask></defs>");out.push_str(&format!("<g mask=\"url(#{mask_id})\">"));}
    let paint=first.color.map_or_else(||"#000000".to_owned(),|c|c.hex());for index in indices{let item=&scene.artworks[*index];if item.shape.operation==ArtworkOperation::Paint{artwork_shape(out,item,*index,&paint,"paint");}}
    if has_erase{out.push_str("</g>");}out.push_str("</g>");
}
pub fn svg(scene:&LayoutSnapshot)->EngineResult<String>{
    let mut out=format!("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{}\" height=\"{}\" viewBox=\"0 0 {} {}\">",number(scene.width),number(scene.height),number(scene.width_em),number(scene.height_em));let mut at=0usize;let mut group=0usize;
    while at<scene.paint_order.len(){match scene.paint_order[at]{
        PaintItem::Glyph{index}=>{let g=&scene.glyphs[index];outline(&mut out,&scene.resources[g.resource].shape.path,g.x,g.y,g.scale,g.color);at+=1;},
        PaintItem::Rule{index}=>{let r=&scene.rules[index];out.push_str(&format!("<rect x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\"",number(r.rect.x_min),number(r.rect.y_min),number(r.rect.width()),number(r.rect.height())));if let Some(color)=r.color{out.push_str(&format!(" fill=\"{}\"",color.hex()));}out.push_str("/>");at+=1;},
        PaintItem::Artwork{index}=>{let first=&scene.artworks[index];let mut indices=Vec::new();while at<scene.paint_order.len(){let PaintItem::Artwork{index}=scene.paint_order[at]else{break};let item=&scene.artworks[index];if item.node_id!=first.node_id||item.asset_sha256!=first.asset_sha256{break;}indices.push(index);at+=1;}artwork_group(&mut out,scene,&indices,group);group+=1;}
    }if out.len()>16*1024*1024{return Err(EngineError::new(ErrorCode::ResourceLimit,"SVG output budget exceeded"));}}
    out.push_str("</svg>");Ok(out)
}
