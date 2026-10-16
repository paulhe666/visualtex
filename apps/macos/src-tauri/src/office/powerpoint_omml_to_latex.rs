//! Native PowerPoint equation -> editable LaTeX, read the same way as the
//! Windows add-in (PowerPointNativeOmmlKernel.ReadNativeMathMl, then
//! MathMlToLatexConverter.ConvertWordEquation): PowerPoint's run colors are
//! carried into Word run properties, then the shared Word equation reader and
//! LaTeX source style apply. Native MathML/text clipboard flavors are not used.
use super::super::omml_latex::{self, dom::{Doc, Name, ParseOptions}};
use super::{parse, Node, A, M};

type Result<T> = std::result::Result<T, String>;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// PowerPointNativeOmmlFormatting.ForWordReverseTransform: copy each run's
/// plain DrawingML RGB color into its Word run properties. Theme colors, alpha
/// and drawing effects need presentation context and are not reinterpreted.
/// Only the characters' own colors count: PowerPoint also leaves colors on a
/// structure's control properties, which would color the whole fraction or
/// accent with its letters.
fn for_word_reverse_transform(omml: &str) -> Result<String> {
    let mut doc = Doc::parse(omml, ParseOptions { keep_whitespace: true, ..ParseOptions::default() })?;
    let root = doc.root.ok_or("The native equation XML has no root element")?;
    for owner in doc.descendants_and_self(root) {
        if !doc.is(owner, M, "r") {
            continue;
        }
        let rgb = doc
            .element(owner, A, "rPr")
            .and_then(|drawing| doc.element(drawing, A, "solidFill"))
            .and_then(|fill| doc.element(fill, A, "srgbClr"));
        let Some(rgb) = rgb else { continue };
        let value = doc.attr(rgb, "val").unwrap_or_default().to_string();
        if !doc.elements(rgb).is_empty() || value.len() != 6 || !value.chars().all(|c| c.is_ascii_hexdigit()) {
            continue;
        }
        let properties = match doc.element(owner, W, "rPr") {
            Some(properties) => properties,
            None => {
                let properties = doc.new_element(Name::new(W, "rPr"));
                match doc.child_nodes(owner).first().copied() {
                    Some(first) => doc.insert_before(first, properties),
                    None => doc.append(owner, properties),
                }
                properties
            }
        };
        for color in doc.elements_named(properties, W, "color") {
            doc.detach(color);
        }
        let color = doc.new_element(Name::new(W, "color"));
        doc.set_attr_ns(color, W, "val", Some(&value.to_ascii_uppercase()));
        doc.append(properties, color);
    }
    Ok(doc.to_xml(root))
}

pub(crate) struct NativeLatex { pub latex:String, pub letter_font:String }
pub(crate) fn from_omml(omml:&str)->Result<NativeLatex> {
    let math_ml = omml_latex::omml_to_mathml(&for_word_reverse_transform(omml)?, true)?;
    let latex = omml_latex::mathml_to_word_equation_latex(&math_ml)?.trim().to_string();
    if latex.is_empty(){return Err("The selected native equation is empty".into());}
    fn fonts(node:&Node,out:&mut Vec<String>){
        if node.name=="a:latin" {if let Some(font)=node.attrs.iter().find(|(k,_)|k=="typeface").map(|(_,v)|v.as_str()){if !out.iter().any(|f|f==font){out.push(font.into());}}}
        for child in &node.children {fonts(child,out);}
    }
    let root=parse(omml)?;
    let mut families=Vec::new();fonts(&root,&mut families);
    let font=families.first().map(String::as_str).unwrap_or("Cambria Math");
    let letter_font=match font {"Times New Roman"=>"times","STIX Two Math"|"STIX Two Text"|"STIX"=>"stix","Palatino"|"Palatino Linotype"=>"palatino","Helvetica"|"Arial"=>"helvetica",_=>"cambria"}.into();
    Ok(NativeLatex{latex,letter_font})
}

#[cfg(test)]
mod tests {
    use super::*;
    fn math(body:&str)->String {format!("<m:oMath xmlns:m=\"{M}\" xmlns:a=\"{A}\" xmlns:w=\"{W}\">{body}</m:oMath>")}
    #[test] fn functions_and_accents_are_structural() {
        let xml=math("<m:func><m:fName><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>sin</m:t></m:r></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func><m:acc><m:accPr><m:chr m:val=\"ˆ\"/></m:accPr><m:e><m:r><m:t>y</m:t></m:r></m:e></m:acc>");
        assert_eq!(from_omml(&xml).unwrap().latex,r"\sin x\hat{y}");
    }
    #[test] fn matrix_and_pre_scripts_are_not_flattened() {
        let xml=math("<m:sPre><m:sub><m:r><m:t>6</m:t></m:r></m:sub><m:sup><m:r><m:t>14</m:t></m:r></m:sup><m:e><m:r><m:t>C</m:t></m:r></m:e></m:sPre>");
        assert_eq!(from_omml(&xml).unwrap().latex,r"{}_6^{14}C");
    }
    /// Real PowerPoint clipboard exports (com.microsoft.Art--Text-ClipFormat),
    /// one file each: `VT_PPT_CLIPBOARDS=<dir> cargo test --lib -- --ignored
    /// real_powerpoint_clipboards --nocapture`.
    #[test]
    #[ignore]
    fn real_powerpoint_clipboards_read_back() {
        let Ok(directory)=std::env::var("VT_PPT_CLIPBOARDS") else {return};
        let mut files:Vec<_>=std::fs::read_dir(directory).unwrap().map(|entry|entry.unwrap().path()).collect();
        files.sort();
        for path in files {
            let omml=super::super::native_equation_from_clipboard(&std::fs::read(&path).unwrap()).unwrap();
            let native=from_omml(&omml).unwrap();
            println!("PPT-LATEX {} {}",path.file_name().unwrap().to_string_lossy(),native.latex);
        }
    }
    #[test] fn plain_rgb_run_colors_reach_the_word_reader() {
        let colored=math("<m:r><a:rPr><a:solidFill><a:srgbClr val=\"ff0000\"/></a:solidFill></a:rPr><m:t>x</m:t></m:r>");
        let prepared=for_word_reverse_transform(&colored).unwrap();
        assert!(prepared.contains("val=\"FF0000\""), "{prepared}");
        // A theme color is left alone.
        let themed=math("<m:r><a:rPr><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill></a:rPr><m:t>x</m:t></m:r>");
        assert!(!for_word_reverse_transform(&themed).unwrap().contains(":color"));
        // A structure's control color does not color its letters.
        let control=math("<m:f><m:fPr><m:ctrlPr><a:rPr><a:solidFill><a:srgbClr val=\"836967\"/></a:solidFill></a:rPr></m:ctrlPr></m:fPr><m:num><m:r><m:t>𝑎</m:t></m:r></m:num><m:den><m:r><m:t>𝑏</m:t></m:r></m:den></m:f>");
        assert_eq!(from_omml(&control).unwrap().latex,r"\frac{a}{b}");
    }
}
