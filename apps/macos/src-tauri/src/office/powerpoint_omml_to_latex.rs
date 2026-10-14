//! Native Office Math -> editable LaTeX. The mathematical tree is authoritative;
//! native MathML/text clipboard flavors are not used to infer structure.
use super::{parse, Node};

type Result<T> = std::result::Result<T, String>;

impl Node {
    fn child(&self, name: &str) -> Option<&Node> { self.children.iter().find(|n| n.name == name) }
    fn value(&self, key: &str) -> Option<&str> { self.attrs.iter().find(|(k,_)| k == key).map(|(_,v)| v.as_str()) }
    fn prop(&self, group: &str, name: &str) -> Option<&str> { self.child(group)?.child(name)?.value("m:val") }
    fn enabled(&self, group: &str, name: &str) -> bool {
        self.child(group).and_then(|p|p.child(name)).is_some_and(|n| !matches!(n.value("m:val"),Some("0"|"false"|"off")))
    }
    fn math_text(&self) -> String {
        if self.name == "m:t" { return self.children.iter().filter_map(|n| n.text.as_deref()).collect(); }
        self.children.iter().filter(|n| n.name.starts_with("m:") && !n.name.ends_with("Pr")).map(Node::math_text).collect()
    }
}

fn escape_text(value: &str) -> String {
    let mut out=String::new();
    for c in value.chars() {
        match c {
            '\\'=>out.push_str("\\textbackslash{}"),
            '{'|'}'|'%'|'#'|'&'|'_'|'$'=>{out.push('\\');out.push(c);},
            '^'=>out.push_str("\\textasciicircum{}"), '~'=>out.push_str("\\textasciitilde{}"),
            _=>out.push(c),
        }
    }
    out
}
fn symbol(value:&str)->String {
    value.chars().map(|c| {
        match c {
            '\u{2009}'=>r"\,".into(), '\u{2005}'=>r"\;".into(), '\u{2003}'=>r"\quad ".into(),
            '\u{2002}'=>r"\enspace ".into(), ' '|'\u{a0}'=>r"\ ".into(),
            '\u{2061}'|'\u{2062}'|'\u{2064}'=>String::new(),
            '^'=>r"\hat{}".into(), '~'=>r"\sim ".into(),
            _=>super::super::omml_batch::native_math_token(&c.to_string()),
        }
    }).collect()
}
fn operator(text:&str)->Option<String> {
    if matches!(text,"sin"|"cos"|"tan"|"cot"|"sec"|"csc"|"sinh"|"cosh"|"tanh"|"coth"|"arcsin"|"arccos"|"arctan"|"log"|"ln"|"exp"|"lim"|"liminf"|"limsup"|"max"|"min"|"sup"|"inf"|"det"|"gcd"|"ker"|"dim"|"Pr"|"arg"|"deg"|"hom") {
        Some(format!("\\{text} "))
    } else { None }
}
fn child_latex(node:&Node,name:&str)->Result<String> {
    node.child(name).ok_or_else(||format!("Native equation is missing {name} in {}",node.name)).and_then(convert)
}
fn sequence(node:&Node)->Result<String> {
    node.children.iter().filter(|n|n.text.is_none() && !n.name.ends_with("Pr")).map(convert).collect()
}
fn delimiter(value:&str,left:bool)->String {
    let side=if left {"left"} else {"right"};
    let value=match value {"{"=>r"\{","}"=>r"\}",""=>".","⟨"|"〈"=>r"\langle","⟩"|"〉"=>r"\rangle","‖"=>r"\Vert","⌊"=>r"\lfloor","⌋"=>r"\rfloor","⌈"=>r"\lceil","⌉"=>r"\rceil",x=>x};
    format!("\\{side}{value} ")
}
fn run(node:&Node)->Result<String> {
    let text=node.math_text();
    let normal=node.enabled("m:rPr","m:nor");
    let style=node.prop("m:rPr","m:sty").unwrap_or("i");
    let script=node.prop("m:rPr","m:scr").unwrap_or("roman");
    let drawing=node.child("a:rPr");
    let mut value=if normal {
        if text.chars().all(|c| c.is_alphanumeric()) && text.chars().count()==1 { format!("\\mathrm{{{}}}",symbol(&text)) }
        else if text.chars().all(|c| !c.is_alphabetic() && !c.is_whitespace()) {symbol(&text)}
        else {format!("\\text{{{}}}",escape_text(&text))}
    } else if style=="p" && script=="roman" && text.chars().any(char::is_alphabetic) && text.chars().all(|c| c as u32 <= 0xffff) {
        operator(&text).unwrap_or_else(||format!("\\mathrm{{{}}}",symbol(&text)))
    } else {symbol(&text)};
    if !normal {
        value=match script {
            "roman"=>value,
            "double-struck"=>format!("\\mathbb{{{value}}}"),
            "script"=>format!("\\mathcal{{{value}}}"),
            "fraktur"=>format!("\\mathfrak{{{value}}}"),
            "sans-serif"=>format!("\\mathsf{{{value}}}"),
            "monospace"=>format!("\\mathtt{{{value}}}"),
            x=>return Err(format!("Unsupported native mathematical script: {x}")),
        };
        if style=="b" || drawing.is_some_and(|n|n.value("b")==Some("1")) {
            value=if style=="bi" {format!("\\boldsymbol{{{value}}}")}else{format!("\\mathbf{{{value}}}")};
        } else if style=="bi" { value=format!("\\boldsymbol{{{value}}}"); }
    }
    if let Some(color)=drawing.and_then(|n|n.child("a:solidFill")).and_then(|n|n.child("a:srgbClr")).and_then(|n|n.value("val")) {
        if color!="000000" && color.len()==6 && color.bytes().all(|c|c.is_ascii_hexdigit()) {value=format!("{{\\color[HTML]{{{color}}}{value}}}");}
    }
    Ok(value)
}
fn accent(node:&Node,group:bool)->Result<String> {
    let props=if group {"m:groupChrPr"}else{"m:accPr"};
    let mark=node.prop(props,"m:chr").unwrap_or(if group {"⏟"}else{"̂"});
    let under=node.prop(props,"m:pos").unwrap_or("bot")=="bot" && group;
    let body=child_latex(node,"m:e")?;
    let command=match mark {
        "ˆ"|"̂"|"^"=>"hat", "˜"|"̃"|"~"=>"tilde", "¯"|"̄"=>if under {"underline"} else {"overline"},
        "˙"|"̇"=>"dot", "¨"|"̈"=>"ddot", "⃛"=>"dddot", "⃜"=>"ddddot",
        "→"|"⃗"=>if group {"overrightarrow"}else{"vec"}, "←"|"⃖"=>"overleftarrow", "↔"|"⃡"=>"overleftrightarrow",
        "⏞"=>"overbrace", "⏟"=>"underbrace", "⏜"=>"overparen", "⏝"=>"underparen",
        "ˇ"|"̌"=>"check", "˘"|"̆"=>"breve", "´"|"́"=>"acute", "`"|"̀"=>"grave",
        _=>return Ok(format!("\\{}{{{}}}{{{body}}}",if under {"underset"}else{"overset"},symbol(mark))),
    };
    Ok(format!("\\{command}{{{body}}}"))
}
fn convert(node:&Node)->Result<String> {
    match node.name.as_str() {
        "m:oMath"|"m:oMathPara"|"m:e"|"m:num"|"m:den"|"m:sub"|"m:sup"|"m:deg"|"m:lim"|"m:fName"=>sequence(node),
        "m:r"=>run(node),
        "m:f"=>{
            let a=child_latex(node,"m:num")?;let b=child_latex(node,"m:den")?;
            Ok(match node.prop("m:fPr","m:type").unwrap_or("bar") {
                "bar"=>format!("\\frac{{{a}}}{{{b}}}"),"noBar"=>format!("{{{a}\\atop {b}}}"),
                "lin"=>format!("{{{a}}}/{{{b}}}"),"skw"=>format!("\\nicefrac{{{a}}}{{{b}}}"),
                x=>return Err(format!("Unknown native fraction type: {x}")),
            })
        }
        "m:rad"=>{
            let body=child_latex(node,"m:e")?;
            let degree=node.child("m:deg").map(convert).transpose()?.unwrap_or_default();
            Ok(if node.enabled("m:radPr","m:degHide")||degree.is_empty() {format!("\\sqrt{{{body}}}")}else{format!("\\sqrt[{degree}]{{{body}}}")})
        }
        "m:sSup"|"m:sSub"|"m:sSubSup"|"m:sPre"=>{
            let body=child_latex(node,"m:e")?;
            let sub=node.child("m:sub").map(convert).transpose()?;
            let sup=node.child("m:sup").map(convert).transpose()?;
            let scripts=format!("{}{}",sub.map(|s|format!("_{{{s}}}")).unwrap_or_default(),sup.map(|s|format!("^{{{s}}}")).unwrap_or_default());
            Ok(if node.name=="m:sPre" {format!("{{}}{scripts}{{{body}}}")}else{format!("{{{body}}}{scripts}")})
        }
        "m:nary"=>{
            let operator=node.prop("m:naryPr","m:chr").unwrap_or("∫");
            let limits=node.prop("m:naryPr","m:limLoc").unwrap_or("subSup");
            let mut result=symbol(operator);
            if limits=="undOvr" {result.push_str("\\limits");}else{result.push_str("\\nolimits");}
            for (name,hide,marker) in [("m:sub","m:subHide","_"),("m:sup","m:supHide","^")] {
                if !node.enabled("m:naryPr",hide) {if let Some(n)=node.child(name){let text=convert(n)?;if !text.is_empty(){result.push_str(&format!("{marker}{{{text}}}"));}}}
            }
            result.push_str(&format!(" {{{}}}",child_latex(node,"m:e")?));Ok(result)
        }
        "m:d"=>{
            let left=node.prop("m:dPr","m:begChr").unwrap_or("(");let right=node.prop("m:dPr","m:endChr").unwrap_or(")");
            let separator=node.prop("m:dPr","m:sepChr").unwrap_or("|");
            let bodies=node.children.iter().filter(|n|n.name=="m:e").map(convert).collect::<Result<Vec<_>>>()?;
            Ok(format!("{}{}{}",delimiter(left,true),bodies.join(&format!("\\middle{} ",if separator==""{"."}else{separator})),delimiter(right,false)))
        }
        "m:m"=>{
            let rows=node.children.iter().filter(|n|n.name=="m:mr").map(|row|row.children.iter().filter(|n|n.name=="m:e").map(convert).collect::<Result<Vec<_>>>().map(|v|v.join(" & "))).collect::<Result<Vec<_>>>()?;
            Ok(format!("\\begin{{matrix}}{}\\end{{matrix}}",rows.join(" \\\\ ")))
        }
        "m:eqArr"=>{
            let mut rows=Vec::new();
            for row in node.children.iter().filter(|n|n.name=="m:e") {
                let mut line=String::new();let mut align=false;
                for child in &row.children {
                    if child.name=="m:r" && !child.enabled("m:rPr","m:nor") && child.math_text().contains('&') {
                        line.push_str(&run(child)?.replace("\\&","&")); align=true; continue;
                    }
                    if child.name=="m:r" && child.enabled("m:rPr","m:aln") {
                        line.push('&');align=true;
                    }
                    if !child.name.ends_with("Pr") && child.text.is_none(){line.push_str(&convert(child)?);}
                }
                if !align {line.insert(0,'&');}
                rows.push(line);
            }
            Ok(format!("\\begin{{aligned}}{}\\end{{aligned}}",rows.join(" \\\\ ")))
        }
        "m:func"=>{
            let name=node.child("m:fName").ok_or("Native function has no name")?;
            let f=operator(&name.math_text()).unwrap_or_else(||format!("\\operatorname{{{}}}",escape_text(&name.math_text())));
            Ok(format!("{f}{{{}}}",child_latex(node,"m:e")?))
        }
        "m:limLow"|"m:limUpp"=>{
            let base=node.child("m:e").ok_or("Native limit has no expression")?;let lower=node.name=="m:limLow";
            let limit=child_latex(node,"m:lim")?;
            if let Some(op)=operator(&base.math_text()) {Ok(format!("{op}{}{{{limit}}}",if lower {"_"}else{"^"}))}
            else{Ok(format!("\\{}{{{limit}}}{{{}}}",if lower{"underset"}else{"overset"},convert(base)?))}
        }
        "m:acc"=>accent(node,false), "m:groupChr"=>accent(node,true),
        "m:bar"=>Ok(format!("\\{}{{{}}}",if node.prop("m:barPr","m:pos").unwrap_or("bot")=="bot"{"underline"}else{"overline"},child_latex(node,"m:e")?)),
        "m:phant"=>{
            let body=child_latex(node,"m:e")?;
            let zero_width=node.enabled("m:phantPr","m:zeroWid");
            let zero_height=node.enabled("m:phantPr","m:zeroAsc")&&node.enabled("m:phantPr","m:zeroDesc");
            let hidden=!node.child("m:phantPr").and_then(|p|p.child("m:show")).is_some_and(|n|!matches!(n.value("m:val"),Some("0"|"false")));
            Ok(if hidden {format!("\\{}{{{body}}}",if zero_width{"vphantom"}else if zero_height{"hphantom"}else{"phantom"})}else if zero_height {format!("\\smash{{{body}}}")}else if zero_width {format!("\\mathrlap{{{body}}}")}else{body})
        }
        "m:box"=>child_latex(node,"m:e"),
        "m:borderBox"=>Ok(format!("\\boxed{{{}}}",child_latex(node,"m:e")?)),
        name if name.ends_with("Pr")=>Ok(String::new()),
        other=>Err(format!("The native equation uses an unhandled mathematical structure ({other}); the original was not changed")),
    }
}

pub(crate) struct NativeLatex { pub latex:String, pub letter_font:String }
pub(crate) fn from_omml(omml:&str)->Result<NativeLatex> {
    let root=parse(omml)?;
    let latex=convert(&root)?.trim().to_string();
    if latex.is_empty(){return Err("The selected native equation is empty".into());}
    fn fonts(node:&Node,out:&mut Vec<String>){
        if node.name=="a:latin" {if let Some(font)=node.value("typeface"){if !out.iter().any(|f|f==font){out.push(font.into());}}}
        for child in &node.children {fonts(child,out);}
    }
    let mut families=Vec::new();fonts(&root,&mut families);
    let font=families.first().map(String::as_str).unwrap_or("Cambria Math");
    let letter_font=match font {"Times New Roman"=>"times","STIX Two Math"|"STIX Two Text"|"STIX"=>"stix","Palatino"|"Palatino Linotype"=>"palatino","Helvetica"|"Arial"=>"helvetica",_=>"cambria"}.into();
    Ok(NativeLatex{latex,letter_font})
}

#[cfg(test)]
mod tests {
    use super::*;
    fn math(body:&str)->String {format!("<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">{body}</m:oMath>")}
    #[test] fn functions_and_accents_are_structural() {
        let xml=math("<m:func><m:fName><m:r><m:t>sin</m:t></m:r></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func><m:acc><m:accPr><m:chr m:val=\"ˆ\"/></m:accPr><m:e><m:r><m:t>y</m:t></m:r></m:e></m:acc>");
        assert_eq!(from_omml(&xml).unwrap().latex,r"\sin {x}\hat{y}");
    }
    #[test] fn matrix_and_pre_scripts_are_not_flattened() {
        let xml=math("<m:sPre><m:sub><m:r><m:t>6</m:t></m:r></m:sub><m:sup><m:r><m:t>14</m:t></m:r></m:sup><m:e><m:r><m:t>C</m:t></m:r></m:e></m:sPre>");
        assert_eq!(from_omml(&xml).unwrap().latex,r"{}_{6}^{14}{C}");
    }
}
