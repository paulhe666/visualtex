//! Bounded vector compiler. Author paths are preserved in the resource snapshot;
//! supported segments become absolute em-space curves. No DOM, SVG markup,
//! system-font lookup, CSS mask, raster tracing or similar-glyph substitution.
use super::*;
use crate::fonts::FontAsset;
use std::f64::consts::PI;
type Point=[f64;2];
fn coord(v:&Value,key:&str)->EngineResult<f64>{let n=number(v,key)?;if n.abs()>100_000.0{Err(bad("Artwork coordinate exceeds the legacy limit"))}else{Ok(n)}}
fn nonnegative(v:&Value,key:&str)->EngineResult<f64>{let n=coord(v,key)?;if n<0.0{Err(bad("Artwork size cannot be negative"))}else{Ok(n)}}
fn line(out:&mut Vec<PathOp>,p:Point){out.push(PathOp::Line(p));}
fn rect(x:f64,y:f64,w:f64,h:f64,reverse:bool)->Vec<PathOp>{
    let points=if reverse{[[x,y],[x,y+h],[x+w,y+h],[x+w,y]]}else{[[x,y],[x+w,y],[x+w,y+h],[x,y+h]]};
    vec![PathOp::Move(points[0]),PathOp::Line(points[1]),PathOp::Line(points[2]),PathOp::Line(points[3]),PathOp::Close]
}
// Cubic approximation of analytic circles/ellipses, with 16 fixed sectors.
// This does not modify imported glyph curves. Same sector count at every zoom.
fn ellipse(x:f64,y:f64,rx:f64,ry:f64,reverse:bool)->Vec<PathOp>{
    let mut out=vec![PathOp::Move([x+rx,y])];let direction=if reverse{-1.0}else{1.0};
    for i in 0..16 {let a=direction*i as f64*PI/8.0;let b=direction*(i+1)as f64*PI/8.0;let k=4.0/3.0*((b-a)/4.0).tan();
        out.push(PathOp::Cubic([x+rx*(a.cos()-k*a.sin()),y+ry*(a.sin()+k*a.cos()),x+rx*(b.cos()+k*b.sin()),y+ry*(b.sin()-k*b.cos()),x+rx*b.cos(),y+ry*b.sin()]));}
    out.push(PathOp::Close);out
}
fn arc_point(cx:f64,cy:f64,rx:f64,ry:f64,phi:f64,t:f64)->Point{let(c,s)=(phi.cos(),phi.sin());[cx+c*rx*t.cos()-s*ry*t.sin(),cy+s*rx*t.cos()+c*ry*t.sin()]}
fn arc_derivative(rx:f64,ry:f64,phi:f64,t:f64)->Point{let(c,s)=(phi.cos(),phi.sin());[-c*rx*t.sin()-s*ry*t.cos(),-s*rx*t.sin()+c*ry*t.cos()]}
fn elliptical_arc(out:&mut Vec<PathOp>,start:Point,mut rx:f64,mut ry:f64,rotation:f64,large:bool,sweep:bool,end:Point)->EngineResult<()> {
    if start==end{return Ok(());}rx=rx.abs();ry=ry.abs();if rx==0.||ry==0.{out.push(PathOp::Line(end));return Ok(());}
    let phi=(rotation%360.)*PI/180.;let(c,s)=(phi.cos(),phi.sin());let dx=(start[0]-end[0])/2.;let dy=(start[1]-end[1])/2.;
    let x1p=c*dx+s*dy;let y1p=-s*dx+c*dy;let lambda=x1p*x1p/(rx*rx)+y1p*y1p/(ry*ry);
    if lambda>1.{let scale=lambda.sqrt();rx*=scale;ry*=scale;}
    let numerator=(rx*rx*ry*ry-rx*rx*y1p*y1p-ry*ry*x1p*x1p).max(0.);let denominator=rx*rx*y1p*y1p+ry*ry*x1p*x1p;
    let coefficient=if denominator==0.{0.}else{(numerator/denominator).sqrt()*if large==sweep{-1.}else{1.}};
    let cxp=coefficient*rx*y1p/ry;let cyp=-coefficient*ry*x1p/rx;
    let cx=c*cxp-s*cyp+(start[0]+end[0])/2.;let cy=s*cxp+c*cyp+(start[1]+end[1])/2.;
    let ux=(x1p-cxp)/rx;let uy=(y1p-cyp)/ry;let vx=(-x1p-cxp)/rx;let vy=(-y1p-cyp)/ry;
    let mut delta=(ux*vy-uy*vx).atan2(ux*vx+uy*vy);if sweep&&delta<0.{delta+=2.*PI;}if !sweep&&delta>0.{delta-=2.*PI;}
    let start_angle=uy.atan2(ux);let segments=(delta.abs()/(PI/2.)).ceil().max(1.)as usize;let step=delta/segments as f64;
    for i in 0..segments {let a=start_angle+i as f64*step;let b=a+step;let alpha=4./3.*((b-a)/4.).tan();let p0=arc_point(cx,cy,rx,ry,phi,a);let p3=arc_point(cx,cy,rx,ry,phi,b);let d0=arc_derivative(rx,ry,phi,a);let d3=arc_derivative(rx,ry,phi,b);
        out.push(PathOp::Cubic([p0[0]+alpha*d0[0],p0[1]+alpha*d0[1],p3[0]-alpha*d3[0],p3[1]-alpha*d3[1],p3[0],p3[1]]));}
    if let Some(PathOp::Cubic(last))=out.last_mut(){last[4]=end[0];last[5]=end[1];}Ok(())
}
fn rounded_rect(x:f64,y:f64,w:f64,h:f64,rx:f64,ry:f64)->EngineResult<Vec<PathOp>>{
    let rx=rx.max(0.).min(w/2.);let ry=ry.max(0.).min(h/2.);if rx==0.||ry==0.{return Ok(rect(x,y,w,h,false));}
    let mut out=vec![PathOp::Move([x+rx,y]),PathOp::Line([x+w-rx,y])];
    elliptical_arc(&mut out,[x+w-rx,y],rx,ry,0.,false,true,[x+w,y+ry])?;out.push(PathOp::Line([x+w,y+h-ry]));
    elliptical_arc(&mut out,[x+w,y+h-ry],rx,ry,0.,false,true,[x+w-rx,y+h])?;out.push(PathOp::Line([x+rx,y+h]));
    elliptical_arc(&mut out,[x+rx,y+h],rx,ry,0.,false,true,[x,y+h-ry])?;out.push(PathOp::Line([x,y+ry]));
    elliptical_arc(&mut out,[x,y+ry],rx,ry,0.,false,true,[x+rx,y])?;out.push(PathOp::Close);Ok(out)
}
#[derive(Clone,Copy)]
struct Matrix([f64;6]);
impl Matrix {
    fn identity()->Self{Self([1.,0.,0.,1.,0.,0.])}
    fn mul(self,b:Self)->Self{let [a,b0,c,d,e,f]=self.0;let [g,h,i,j,k,l]=b.0;Self([a*g+c*h,b0*g+d*h,a*i+c*j,b0*i+d*j,a*k+c*l+e,b0*k+d*l+f])}
    fn translate(x:f64,y:f64)->Self{Self([1.,0.,0.,1.,x,y])}
    fn point(self,[x,y]:Point)->EngineResult<Point>{let [a,b,c,d,e,f]=self.0;let p=[a*x+c*y+e,b*x+d*y+f];
        if p.iter().any(|v|!v.is_finite()||v.abs()>10_000_000.0){return Err(limit("Transformed artwork exceeds the vector coordinate budget"));}Ok(p)}
    fn em_transform(self,ascent:f64)->[f64;6]{let [a,b,c,d,e,f]=self.0;[a,b,c,d,c*ascent+e/1000.,(d-1.)*ascent+f/1000.]}
    fn max_linear_scale(self)->f64{let[a,b,c,d,_,_]=self.0;let trace=a*a+b*b+c*c+d*d;let det=a*d-b*c;((trace+(trace*trace-4.*det*det).max(0.).sqrt())/2.).sqrt()}
}
fn transforms(value:Option<&Value>)->EngineResult<(Matrix,Matrix)>{
    let Some(v)=value else{return Ok((Matrix::identity(),Matrix::identity()))};if !v.is_object(){return Err(bad("Artwork transform must be an object"));}
    let get=|key,default|optional_number(v,key,default);
    let(tx,ty,ox,oy)=(get("translateX",0.)?,get("translateY",0.)?,get("originX",0.)?,get("originY",0.)?);
    let(sx,sy)=(get("scaleX",1.)?,get("scaleY",1.)?);if sx==0.||sy==0.{return Err(bad("Artwork transform cannot have a zero scale"));}
    let angle=get("rotateDeg",0.)?*PI/180.;let skewx=(get("skewXDeg",0.)?*PI/180.).tan();let skewy=(get("skewYDeg",0.)?*PI/180.).tan();
    let inner=if let Some(array)=v.get("matrix") {let a=array.as_array().filter(|a|a.len()==6).ok_or_else(||bad("Artwork matrix needs six values"))?;let mut n=[0.;6];for(i,v)in a.iter().enumerate(){n[i]=v.as_f64().filter(|n|n.is_finite()&&n.abs()<=100_000.).ok_or_else(||bad("Invalid artwork matrix coordinate"))?;}if (n[0]*n[3]-n[1]*n[2]).abs()<1e-7{return Err(bad("Artwork matrix is singular"));}Matrix(n)}else{Matrix::identity()};
    let outer=Matrix::translate(tx,ty).mul(Matrix::translate(ox,oy)).mul(Matrix([angle.cos(),angle.sin(),-angle.sin(),angle.cos(),0.,0.]))
        .mul(Matrix([1.,0.,skewx,1.,0.,0.])).mul(Matrix([1.,skewy,0.,1.,0.,0.])).mul(Matrix([sx,0.,0.,sy,0.,0.])).mul(Matrix::translate(-ox,-oy));
    let combined=outer.mul(inner);if outer.0.iter().chain(inner.0.iter()).chain(combined.0.iter()).any(|v|!v.is_finite()||v.abs()>1e8){return Err(bad("Artwork transform is non-finite or excessive"));}Ok((outer,inner))
}
fn clip_rect(value:Option<&Value>,ascent:f64)->EngineResult<Option<Rect>>{
    let Some(v)=value else{return Ok(None)};if !v.is_object(){return Err(bad("Artwork clip rectangle must be an object"));}
    let(x,y,w,h)=(coord(v,"x")?,coord(v,"y")?,coord(v,"width")?,coord(v,"height")?);if w<=0.||h<=0.{return Err(bad("Artwork clip rectangle must have positive size"));}
    Ok(Some(Rect::new(x/1000.,y/1000.-ascent,(x+w)/1000.,(y+h)/1000.-ascent)))
}
fn transformed_rect(matrix:[f64;6],rect:Rect)->Rect{
    let[a,b,c,d,e,f]=matrix;let p=|x:f64,y:f64|[a*x+c*y+e,b*x+d*y+f];let points=[p(rect.x_min,rect.y_min),p(rect.x_max,rect.y_min),p(rect.x_max,rect.y_max),p(rect.x_min,rect.y_max)];
    Rect::new(points.iter().map(|p|p[0]).fold(f64::INFINITY,f64::min),points.iter().map(|p|p[1]).fold(f64::INFINITY,f64::min),points.iter().map(|p|p[0]).fold(f64::NEG_INFINITY,f64::max),points.iter().map(|p|p[1]).fold(f64::NEG_INFINITY,f64::max))
}
fn intersect(a:Rect,b:Rect)->Rect{let x0=a.x_min.max(b.x_min);let y0=a.y_min.max(b.y_min);let x1=a.x_max.min(b.x_max);let y1=a.y_max.min(b.y_max);if x1<x0||y1<y0{Rect::new(x0,y0,x0,y0)}else{Rect::new(x0,y0,x1,y1)}}
fn finish(path:Vec<PathOp>,transform:[f64;6],inner_transform:[f64;6],clip:Option<Rect>,text_source:Option<ArtworkTextSource>,outline_font_id:Option<String>,operation:ArtworkOperation,fill:bool,stroke_width:f64,cap:ArtworkLineCap,join:ArtworkLineJoin)->EngineResult<ArtworkPath>{
    let combined=Matrix(transform).mul(Matrix(inner_transform));let mut bounds:Option<Rect>=None;
    let mut include=|p:Point|->EngineResult<()> {let p=combined.point(p)?;let r=Rect::new(p[0],p[1],p[0],p[1]);bounds=Some(bounds.map_or(r,|b|b.union(r)));Ok(())};
    for op in &path{match op{PathOp::Move(p)|PathOp::Line(p)=>include(*p)?,PathOp::Quad(p)=>{include([p[0],p[1]])?;include([p[2],p[3]])?;},PathOp::Cubic(p)=>{include([p[0],p[1]])?;include([p[2],p[3]])?;include([p[4],p[5]])?;},PathOp::Close=>{}}}
    let mut ink=bounds.unwrap_or(Rect::new(0.,0.,0.,0.));if stroke_width>0.{let factor=if join==ArtworkLineJoin::Miter{2.0}else if cap==ArtworkLineCap::Square{std::f64::consts::FRAC_1_SQRT_2}else{0.5};let pad=stroke_width*factor*combined.max_linear_scale();ink=Rect::new(ink.x_min-pad,ink.y_min-pad,ink.x_max+pad,ink.y_max+pad);}
    if let Some(clip)=clip{ink=intersect(ink,transformed_rect(transform,clip));}
    Ok(ArtworkPath{path,transform,inner_transform,ink,clip,text_source,outline_font_id,operation,fill,stroke_width,line_cap:cap,line_join:join})
}
fn styled(path:Vec<PathOp>,outer:Matrix,inner:Matrix,ascent:f64,clip:Option<Rect>,operation:ArtworkOperation,fill:bool,stroke:f64,cap:ArtworkLineCap,join:ArtworkLineJoin)->EngineResult<ArtworkPath>{
    let q=|n:f64|(n*1e9).round()/1e9;let mut result=Vec::with_capacity(path.len());let point=|p:Point|[q(p[0]/1000.),q(p[1]/1000.-ascent)];
    for op in path {result.push(match op{PathOp::Move(p)=>PathOp::Move(point(p)),PathOp::Line(p)=>PathOp::Line(point(p)),PathOp::Quad(p)=>{let a=point([p[0],p[1]]);let b=point([p[2],p[3]]);PathOp::Quad([a[0],a[1],b[0],b[1]])},PathOp::Cubic(p)=>{let a=point([p[0],p[1]]);let b=point([p[2],p[3]]);let c=point([p[4],p[5]]);PathOp::Cubic([a[0],a[1],b[0],b[1],c[0],c[1]])},PathOp::Close=>PathOp::Close});}
    finish(result,outer.em_transform(ascent),inner.em_transform(ascent),clip,None,None,operation,fill,stroke/1000.,cap,join)
}
fn text_source(shape:&Value)->EngineResult<ArtworkTextSource>{
    let text=string(shape,"text")?;if text.is_empty()||text.chars().count()>16||text.chars().any(char::is_control){return Err(bad("Invalid custom text artwork content"));}
    let font_family=string(shape,"fontFamily")?.trim();if font_family.is_empty()||font_family.len()>128||font_family.chars().any(|c|c.is_control()||matches!(c,'\"'|'\''|'<'|'>')){return Err(bad("Invalid custom text artwork font family"));}
    let font_size=number(shape,"fontSize")?;if !(20.0..=8000.0).contains(&font_size){return Err(bad("Invalid custom text artwork font size"));}
    let font_style=match shape.get("fontStyle").and_then(Value::as_str).unwrap_or("normal"){"normal"=>ArtworkFontStyle::Normal,"italic"=>ArtworkFontStyle::Italic,_=>return Err(bad("Invalid custom text artwork font style"))};
    let weight=optional_number(shape,"fontWeight",400.)?;if weight.fract()!=0.||!(100.0..=900.0).contains(&weight){return Err(bad("Invalid custom text artwork font weight"));}
    Ok(ArtworkTextSource{text:text.to_owned(),x:coord(shape,"x")?,y:coord(shape,"y")?,font_family:font_family.to_owned(),font_size,font_style,font_weight:weight as u16})
}
pub(super) fn compile(shape:&Value,ascent:f64)->EngineResult<Vec<ArtworkPath>>{
    if !shape.is_object(){return Err(bad("Artwork shape must be an object"));}
    let kind=string(shape,"kind")?;let(outer,inner)=transforms(shape.get("transform"))?;let clip=clip_rect(shape.get("clipRect"),ascent)?;
    let operation=match shape.get("operation"){None=>ArtworkOperation::Paint,Some(Value::String(op))if op=="erase"=>ArtworkOperation::Erase,Some(_)=>return Err(bad("Unknown artwork operation"))};
    let fill=match shape.get("fill"){None=>kind!="line",Some(Value::Bool(b))=>*b,_=>return Err(bad("Invalid artwork fill value"))};
    let stroke=optional_number(shape,"strokeWidth",if kind=="line"{50.}else{0.})?;
    if !(0.0..=100_000.).contains(&stroke){return Err(bad("Invalid artwork stroke width"));}
    let cap=match shape.get("lineCap").and_then(Value::as_str).unwrap_or("butt"){
        "butt"=>ArtworkLineCap::Butt,"round"=>ArtworkLineCap::Round,"square"=>ArtworkLineCap::Square,_=>return Err(bad("Unsupported line cap value")),
    };
    let join=match shape.get("lineJoin").and_then(Value::as_str).unwrap_or("miter"){
        "miter"=>ArtworkLineJoin::Miter,"round"=>ArtworkLineJoin::Round,"bevel"=>ArtworkLineJoin::Bevel,_=>return Err(bad("Unsupported line join value")),
    };
    if kind=="text"{let source=text_source(shape)?;return Ok(vec![ArtworkPath{path:Vec::new(),transform:outer.em_transform(ascent),inner_transform:inner.em_transform(ascent),ink:Rect::new(0.,0.,0.,0.),clip,text_source:Some(source),outline_font_id:None,operation,fill,stroke_width:stroke/1000.,line_cap:cap,line_join:join}]);}
    let mut paths=Vec::new();
    match kind{
        "path"=>{if fill||stroke>0.{paths.push(parse_path(string(shape,"d")?)?);}},
        "line"=>{let(x1,y1,x2,y2)=(coord(shape,"x1")?,coord(shape,"y1")?,coord(shape,"x2")?,coord(shape,"y2")?);if stroke>0.{paths.push(vec![PathOp::Move([x1,y1]),PathOp::Line([x2,y2])]);}},
        "circle"=>{let(x,y,r)=(coord(shape,"cx")?,coord(shape,"cy")?,nonnegative(shape,"r")?);if (fill||stroke>0.)&&r>0.{paths.push(ellipse(x,y,r,r,false));}},
        "ellipse"=>{let(x,y,rx,ry)=(coord(shape,"cx")?,coord(shape,"cy")?,nonnegative(shape,"rx")?,nonnegative(shape,"ry")?);if (fill||stroke>0.)&&rx>0.&&ry>0.{paths.push(ellipse(x,y,rx,ry,false));}},
        "rect"=>{let(x,y,w,h)=(coord(shape,"x")?,coord(shape,"y")?,nonnegative(shape,"width")?,nonnegative(shape,"height")?);let mut rx=optional_number(shape,"rx",0.)?;let mut ry=optional_number(shape,"ry",0.)?;if rx<0.||ry<0.{return Err(bad("Rounded rectangle radii cannot be negative"));}if rx==0.&&ry>0.{rx=ry;}if ry==0.&&rx>0.{ry=rx;}if (fill||stroke>0.)&&w>0.&&h>0.{paths.push(rounded_rect(x,y,w,h,rx,ry)?);}},
        "polygon"=>{let points=shape["points"].as_array().filter(|p|(2..=256).contains(&p.len())).ok_or_else(||bad("Invalid polygon vertex count"))?;let mut path=Vec::new();
            for (i,p) in points.iter().enumerate(){let p=p.as_array().filter(|p|p.len()==2).ok_or_else(||bad("Invalid polygon vertex"))?;let mut xy=[0.;2];for k in 0..2{xy[k]=p[k].as_f64().filter(|n|n.is_finite()&&n.abs()<=100_000.).ok_or_else(||bad("Invalid polygon coordinate"))?;}path.push(if i==0{PathOp::Move(xy)}else{PathOp::Line(xy)});}path.push(PathOp::Close);if fill||stroke>0.{paths.push(path);}},
        _=>return Err(bad("Unknown custom artwork shape kind")),
    }
    paths.into_iter().filter(|p|!p.is_empty()).map(|p|styled(p,outer,inner,ascent,clip,operation,fill,stroke,cap,join)).collect()
}
pub(crate) fn materialize_text(template:&ArtworkPath,font:&FontAsset,ascent:f64)->EngineResult<ArtworkPath>{
    let source=template.text_source.as_ref().ok_or_else(||bad("Custom artwork is not an unresolved text shape"))?;let run=font.shape_text(&source.text)?;let scale=source.font_size/1000.;let origin=[source.x/1000.,source.y/1000.-ascent];let mut path=Vec::new();
    let map=|p:Point,x:f64,y:f64|[origin[0]+(x+p[0])*scale,origin[1]+(y+p[1])*scale];
    for glyph in run.glyphs{for op in &glyph.shape.path{if path.len()>=MAX_ARTWORK_OPS{return Err(limit("Custom text outline exceeds the compiled path budget"));}path.push(match op{PathOp::Move(p)=>PathOp::Move(map(*p,glyph.x,glyph.y)),PathOp::Line(p)=>PathOp::Line(map(*p,glyph.x,glyph.y)),PathOp::Quad(p)=>{let a=map([p[0],p[1]],glyph.x,glyph.y);let b=map([p[2],p[3]],glyph.x,glyph.y);PathOp::Quad([a[0],a[1],b[0],b[1]])},PathOp::Cubic(p)=>{let a=map([p[0],p[1]],glyph.x,glyph.y);let b=map([p[2],p[3]],glyph.x,glyph.y);let c=map([p[4],p[5]],glyph.x,glyph.y);PathOp::Cubic([a[0],a[1],b[0],b[1],c[0],c[1]])},PathOp::Close=>PathOp::Close});}}
    if path.is_empty(){return Err(bad("Bound custom text font produced no vector outline"));}
    finish(path,template.transform,template.inner_transform,template.clip,None,Some(font.identity.id.clone()),template.operation,template.fill,template.stroke_width,template.line_cap,template.line_join)
}
struct Scanner<'a>{bytes:&'a[u8],at:usize}
impl Scanner<'_>{
    fn skip(&mut self){while self.bytes.get(self.at).is_some_and(|b|b.is_ascii_whitespace()||*b==b','){self.at+=1;}}
    fn number(&mut self)->EngineResult<f64>{self.skip();let start=self.at;if self.bytes.get(self.at).is_some_and(|b|matches!(b,b'+'|b'-')){self.at+=1;}
        let mut digits=0;while self.bytes.get(self.at).is_some_and(u8::is_ascii_digit){self.at+=1;digits+=1;}
        if self.bytes.get(self.at)==Some(&b'.'){self.at+=1;while self.bytes.get(self.at).is_some_and(u8::is_ascii_digit){self.at+=1;digits+=1;}}
        if digits==0{return Err(bad("Malformed SVG path number"));}
        if self.bytes.get(self.at).is_some_and(|b|matches!(b,b'e'|b'E')){self.at+=1;if self.bytes.get(self.at).is_some_and(|b|matches!(b,b'+'|b'-')){self.at+=1;}let exponent=self.at;while self.bytes.get(self.at).is_some_and(u8::is_ascii_digit){self.at+=1;}if self.at==exponent{return Err(bad("Malformed SVG path exponent"));}}
        let n=std::str::from_utf8(&self.bytes[start..self.at]).ok().and_then(|s|s.parse::<f64>().ok()).filter(|n|n.is_finite()&&n.abs()<=100_000.).ok_or_else(||bad("SVG path number exceeds the coordinate budget"))?;Ok(n)
    }
    fn point(&mut self,current:Point,relative:bool)->EngineResult<Point>{let p=[self.number()?,self.number()?];let p=if relative{[p[0]+current[0],p[1]+current[1]]}else{p};if p.iter().any(|n|n.abs()>100_000.){return Err(bad("Relative SVG path exceeds the coordinate budget"));}Ok(p)}
    fn flag(&mut self)->EngineResult<bool>{let value=self.number()?;if value==0.{Ok(false)}else if value==1.{Ok(true)}else{Err(bad("SVG arc flags must be 0 or 1"))}}
}
fn parse_path(source:&str)->EngineResult<Vec<PathOp>>{
    if source.is_empty()||source.len()>24_000{return Err(limit("SVG path source exceeds the legacy path budget"));}
    let mut s=Scanner{bytes:source.as_bytes(),at:0};let mut out=Vec::new();let mut current=[0.;2];let mut start=current;let mut command=None;let mut previous=0u8;let mut control=current;
    loop{s.skip();let Some(&next)=s.bytes.get(s.at)else{break};
        if next.is_ascii_alphabetic(){command=Some(next);s.at+=1;}
        let c=command.ok_or_else(||bad("SVG path requires an initial move command"))?;let upper=c.to_ascii_uppercase();let relative=c.is_ascii_lowercase();
        if out.is_empty()&&upper!=b'M'{return Err(bad("SVG path must begin with moveto"));}
        match upper {
            b'M'=>{current=s.point(current,relative)?;start=current;out.push(PathOp::Move(current));command=Some(if relative{b'l'}else{b'L'});},
            b'L'=>{current=s.point(current,relative)?;line(&mut out,current);},
            b'H'=>{let x=s.number()?;current[0]=x+if relative{current[0]}else{0.};line(&mut out,current);},
            b'V'=>{let y=s.number()?;current[1]=y+if relative{current[1]}else{0.};line(&mut out,current);},
            b'C'=>{let a=s.point(current,relative)?;let b=s.point(current,relative)?;let end=s.point(current,relative)?;out.push(PathOp::Cubic([a[0],a[1],b[0],b[1],end[0],end[1]]));control=b;current=end;},
            b'S'=>{let a=if matches!(previous,b'C'|b'S'){[2.*current[0]-control[0],2.*current[1]-control[1]]}else{current};let b=s.point(current,relative)?;let end=s.point(current,relative)?;out.push(PathOp::Cubic([a[0],a[1],b[0],b[1],end[0],end[1]]));control=b;current=end;},
            b'Q'=>{let a=s.point(current,relative)?;let end=s.point(current,relative)?;out.push(PathOp::Quad([a[0],a[1],end[0],end[1]]));control=a;current=end;},
            b'T'=>{let a=if matches!(previous,b'Q'|b'T'){[2.*current[0]-control[0],2.*current[1]-control[1]]}else{current};let end=s.point(current,relative)?;out.push(PathOp::Quad([a[0],a[1],end[0],end[1]]));control=a;current=end;},
            b'Z'=>{out.push(PathOp::Close);current=start;command=None;},
            b'A'=>{let rx=s.number()?.abs();let ry=s.number()?.abs();let rotation=s.number()?;let large=s.flag()?;let sweep=s.flag()?;let end=s.point(current,relative)?;elliptical_arc(&mut out,current,rx,ry,rotation,large,sweep,end)?;current=end;},
            _=>return Err(bad("Unsupported or unsafe SVG path token")),
        }
        if current.iter().any(|v|v.abs()>100_000.){return Err(bad("SVG path exceeds the coordinate budget"));}
        if out.len()>16_384{return Err(limit("SVG path exceeds the segment budget"));}previous=upper;
    }
    if out.is_empty(){return Err(bad("Empty SVG path"));}Ok(out)
}
