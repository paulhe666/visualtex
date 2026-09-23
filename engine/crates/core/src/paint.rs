//! Bounded color semantics, independent of CSS, DOM and TeX execution.
use serde::{Deserialize,Serialize};
use crate::error::{EngineError,EngineResult,ErrorCode};

pub const MAX_COLOR_SOURCE:usize=1024;
#[derive(Debug,Clone,Copy,PartialEq,Eq,Serialize,Deserialize)]
#[serde(deny_unknown_fields)]
pub struct RgbColor {pub r:u8,pub g:u8,pub b:u8}
impl RgbColor {pub fn hex(self)->String{format!("#{:02x}{:02x}{:02x}",self.r,self.g,self.b)}}
#[derive(Debug,Clone,PartialEq,Eq,Serialize)]
#[serde(rename_all="camelCase")]
pub struct ColorSpec {
    pub before_model:String,
    pub model:Option<String>,
    pub before_value:String,
    pub value:String,
    pub color:RgbColor,
}
impl ColorSpec {
    pub fn authored(before_model:String,model:Option<String>,before_value:String,value:String)->EngineResult<Self>{
        let color=resolve_color(model.as_deref(),&value)?;
        let spec=Self{before_model,model,before_value,value,color};spec.validate()?;Ok(spec)
    }
    pub fn html(color:RgbColor)->Self{Self{before_model:String::new(),model:Some("HTML".into()),before_value:String::new(),
        value:format!("{:02X}{:02X}{:02X}",color.r,color.g,color.b),color}}
    pub fn validate(&self)->EngineResult<()> {
        if self.before_model.len()+self.before_value.len()+self.value.len()+self.model.as_ref().map_or(0,String::len)>MAX_COLOR_SOURCE {
            return Err(EngineError::new(ErrorCode::ResourceLimit,"Color source budget exceeded"));
        }
        if self.model.is_none()&&!self.before_model.is_empty()
            ||!without_comments(&self.before_model).trim().is_empty()||!without_comments(&self.before_value).trim().is_empty()
            ||resolve_color(self.model.as_deref(),&self.value)?!=self.color {
            return Err(EngineError::new(ErrorCode::InvalidInput,"Color source and resolved paint disagree"));
        }Ok(())
    }
    pub fn emit(&self,lossless:bool,out:&mut String){
        if let Some(model)=&self.model {
            if lossless {out.push_str(&self.before_model);}
            out.push('[');out.push_str(if lossless{model}else{model.trim()});out.push(']');
        }
        if lossless {out.push_str(&self.before_value);}
        out.push('{');if lossless{out.push_str(&self.value);}else{out.push_str(without_comments(&self.value).trim());}out.push('}');
    }
}
fn unsupported()->EngineError{EngineError::new(ErrorCode::UnsupportedSyntax,"Color model or value is unsupported or invalid")}
fn without_comments(source:&str)->String{
    let mut output=String::new();let mut comment=false;
    for c in source.chars(){if c=='\n'||c=='\r'{comment=false;output.push(' ');}else if c=='%'{comment=true;}else if !comment{output.push(c);}}
    output
}
fn decimal(source:&str,max:f64)->EngineResult<f64>{
    let source=source.trim();
    if source.is_empty()||!source.chars().all(|c|c.is_ascii_digit()||matches!(c,'.'|'+'|'-')){return Err(unsupported());}
    let n=source.parse::<f64>().map_err(|_|unsupported())?;
    if !n.is_finite()||!(0.0..=max).contains(&n){return Err(unsupported());}Ok(n)
}
fn named(name:&str)->Option<[f64;3]>{
    Some(match name.trim(){
        "black"=>[0.,0.,0.],"white"=>[1.,1.,1.],"red"=>[1.,0.,0.],"green"=>[0.,1.,0.],"blue"=>[0.,0.,1.],
        "cyan"=>[0.,1.,1.],"magenta"=>[1.,0.,1.],"yellow"=>[1.,1.,0.],
        "gray"=>[0.5,0.5,0.5],"darkgray"=>[0.25,0.25,0.25],"lightgray"=>[0.75,0.75,0.75],
        "brown"=>[0.75,0.5,0.25],"lime"=>[0.75,1.,0.],"olive"=>[0.5,0.5,0.],"orange"=>[1.,0.5,0.],
        "pink"=>[1.,0.75,0.75],"purple"=>[0.75,0.,0.25],"teal"=>[0.,0.5,0.5],"violet"=>[0.5,0.,0.5],_=>return None,
    })
}
fn named_expression(value:&str)->EngineResult<[f64;3]>{
    let parts:Vec<_>=value.split('!').collect();
    if parts.len()>33{return Err(EngineError::new(ErrorCode::ResourceLimit,"Color mixing budget exceeded"));}
    let mut color=named(parts[0]).ok_or_else(unsupported)?;let mut index=1;
    while index<parts.len(){
        let weight=decimal(parts[index],100.0)?/100.;index+=1;
        let other=if index<parts.len(){let c=named(parts[index]).ok_or_else(unsupported)?;index+=1;c}else{[1.,1.,1.]};
        color=std::array::from_fn(|i|weight*color[i]+(1.-weight)*other[i]);
    }Ok(color)
}
pub fn resolve_color(model:Option<&str>,value:&str)->EngineResult<RgbColor>{
    if value.len()+model.map_or(0,str::len)>MAX_COLOR_SOURCE {return Err(EngineError::new(ErrorCode::ResourceLimit,"Color source budget exceeded"));}
    let model=without_comments(model.unwrap_or("named"));let value=without_comments(value);let value=value.trim();
    let channels=match model.trim(){
        "named"=>named_expression(value)?,
        "HTML"=>{
            if value.len()!=6||!value.bytes().all(|c|c.is_ascii_hexdigit()){return Err(unsupported());}
            return Ok(RgbColor{r:u8::from_str_radix(&value[0..2],16).map_err(|_|unsupported())?,
                g:u8::from_str_radix(&value[2..4],16).map_err(|_|unsupported())?,b:u8::from_str_radix(&value[4..6],16).map_err(|_|unsupported())?});
        }
        "rgb"|"RGB"=>{
            let values:Vec<_>=value.split(',').collect();if values.len()!=3{return Err(unsupported());}
            let upper=if model.trim()=="RGB"{255.}else{1.};
            let mut channels=[0.;3];for(i,value)in values.iter().enumerate(){
                let value=decimal(value,upper)?;if upper==255.&&value.fract()!=0.0{return Err(unsupported());}channels[i]=value/upper;
            }channels
        }
        "gray"=>{let n=decimal(value,1.)?;[n,n,n]}
        _=>return Err(unsupported()),
    };
    Ok(RgbColor{r:(channels[0]*255.).round()as u8,g:(channels[1]*255.).round()as u8,b:(channels[2]*255.).round()as u8})
}
#[cfg(test)]mod tests{
    use super::*;
    #[test]fn bounded_models_resolve_to_safe_rgb_not_css_strings(){
        for(model,value,expected)in[(None,"red","#ff0000"),(Some("HTML"),"1a2B3c","#1a2b3c"),
            (Some("RGB"),"0, 128, 255","#0080ff"),(Some("rgb"),"0,.5,1","#0080ff"),(Some("gray"),".5","#808080"),
            (None,"red!50!blue","#800080"),(None,"red!50","#ff8080")]{assert_eq!(resolve_color(model,value).unwrap().hex(),expected);}
    }
    #[test]fn malformed_nonfinite_and_injection_values_fail_closed(){
        for(model,value)in[(None,"url(https://example.invalid)"),(None,"red;fill:blue"),(None,"currentColor"),(None,"red!101"),
            (Some("HTML"),"#123456"),(Some("HTML"),"123\"45"),(Some("rgb"),"NaN,0,0"),(Some("rgb"),"-1,0,0"),
            (Some("rgb"),"1e-1,0,0"),(Some("RGB"),"1.5,0,0"),(Some("gray"),"Infinity"),(Some("cmyk"),"0,0,0,0")]{assert!(resolve_color(model,value).is_err(),"{model:?} {value}");}
    }
}
