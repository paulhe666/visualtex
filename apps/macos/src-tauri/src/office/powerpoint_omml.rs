//! PowerPoint's native mathematical TEXT clipboard interchange.
//! The DrawingML envelope is consumed by TextRange.Paste, not inserted as a
//! shape/group. Only a14:m / m:oMath remains in the destination text paragraph.
//! No source presentation, Word process, image fallback or OLE pointer is used.
use quick_xml::{events::Event, Reader};
use std::io::{Cursor, Read};
#[path = "powerpoint_omml_to_latex.rs"]
pub(super) mod to_latex;

const M: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A14: &str = "http://schemas.microsoft.com/office/drawing/2010/main";
const LC: &str = "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas";
const REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

#[derive(Debug, Default, Clone)]
struct Node {
    name: String,
    attrs: Vec<(String, String)>,
    children: Vec<Node>,
    text: Option<String>,
}
fn escape(s: &str) -> String { quick_xml::escape::escape(s).into_owned() }
fn parse_node(xml: &str) -> Result<Node, String> {
    let mut reader = Reader::from_str(xml);
    let mut stack: Vec<Node> = vec![Node::default()];
    loop {
        match reader.read_event().map_err(|e| format!("Invalid Office Math XML: {e}"))? {
            event @ (Event::Start(_) | Event::Empty(_)) => {
                let empty = matches!(&event, Event::Empty(_));
                let e = match event { Event::Start(e) | Event::Empty(e) => e, _ => unreachable!() };
                let name = String::from_utf8(e.name().as_ref().to_vec()).map_err(|_| "Invalid XML name")?;
                let mut attrs = Vec::new();
                for attr in e.attributes() {
                    let attr = attr.map_err(|e| e.to_string())?;
                    let key = String::from_utf8(attr.key.as_ref().to_vec()).map_err(|_| "Invalid XML attribute")?;
                    let value = attr.decode_and_unescape_value(reader.decoder()).map_err(|e| e.to_string())?.into_owned();
                    if key == "xmlns:m" && value != M { return Err("Invalid Office Math namespace".into()); }
                    attrs.push((key, value));
                }
                let node = Node { name, attrs, ..Node::default() };
                if empty { stack.last_mut().unwrap().children.push(node); } else { stack.push(node); }
                if stack.len() > 256 { return Err("Office Math XML is nested too deeply".into()); }
            }
            Event::End(_) => {
                if stack.len() < 2 { return Err("Unbalanced Office Math XML".into()); }
                let node = stack.pop().unwrap();
                stack.last_mut().unwrap().children.push(node);
            }
            Event::Text(e) => {
                let text = e.decode().map_err(|e| e.to_string())?;
                if stack.len() > 1 { stack.last_mut().unwrap().children.push(Node { text: Some(text.into_owned()), ..Node::default() }); }
                else if !text.trim().is_empty() { return Err("Unexpected text outside Office Math".into()); }
            }
            Event::CData(e) => stack.last_mut().unwrap().children.push(Node { text: Some(e.decode().map_err(|e| e.to_string())?.into_owned()), ..Node::default() }),
            Event::GeneralRef(e) => {
                let reference = format!("&{};", e.decode().map_err(|e| e.to_string())?);
                let text = quick_xml::escape::unescape(&reference).map_err(|e| e.to_string())?.into_owned();
                stack.last_mut().unwrap().children.push(Node { text: Some(text), ..Node::default() });
            }
            Event::DocType(_) | Event::PI(_) => return Err("Office Math XML cannot contain declarations or processing instructions".into()),
            Event::Eof => break,
            _ => {}
        }
    }
    if stack.len() != 1 || stack[0].children.len() != 1 { return Err("Expected exactly one Office Math root".into()); }
    Ok(stack.pop().unwrap().children.remove(0))
}
fn parse(xml: &str) -> Result<Node, String> {
    let root = parse_node(xml)?;
    if root.name != "m:oMath" || !root.attrs.iter().any(|(k,v)| k == "xmlns:m" && v == M) {
        return Err("Expected a namespaced m:oMath root".into());
    }
    Ok(root)
}
fn property<'a>(node: &'a Node, name: &str, key: &str) -> Option<&'a str> {
    node.children.iter().find(|n| n.name == name)?.attrs.iter().find(|(k,_)| k == key).map(|(_,v)| v.as_str())
}
fn run_properties(word: Option<&Node>, size: i64, language: Option<&str>) -> String {
    let mut flags = String::new();
    if let Some(word) = word {
        for (w,a) in [("w:b","b"),("w:i","i")] {
            if let Some(n) = word.children.iter().find(|n| n.name == w) {
                let enabled = !n.attrs.iter().any(|(k,v)| k == "w:val" && matches!(v.as_str(), "0"|"false"|"off"));
                flags.push_str(&format!(" {a}=\"{}\"", if enabled {1} else {0}));
            }
        }
    }
    let mut xml = match language {
        Some(language) => format!("<a:rPr lang=\"{language}\" altLang=\"{language}\" sz=\"{size}\"{flags}>"),
        None => format!("<a:rPr sz=\"{size}\"{flags}>"),
    };
    if let Some(color) = word.and_then(|n| property(n,"w:color","w:val")) {
        if color.len() == 6 && color.bytes().all(|b| b.is_ascii_hexdigit()) {
            xml.push_str(&format!("<a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>"));
        }
    }
    for (key,tag,fallback) in [("w:ascii","a:latin",Some("Cambria Math")),("w:eastAsia","a:ea",None),("w:cs","a:cs",None)] {
        if let Some(font) = word.and_then(|n| property(n,"w:rFonts",key)).or(fallback) {
            xml.push_str(&format!("<{tag} typeface=\"{}\"/>", escape(font)));
        }
    }
    xml.push_str("</a:rPr>"); xml
}
/// OMML properties of on/off type (ST_OnOff). PowerPoint's native clipboard
/// importer drops their numeric values (a radical's degree, an operator's empty
/// limit and a matrix's placeholders reappear), so they are written in its own
/// canonical on/off form.
const ON_OFF: [&str; 30] = [
    "m:aln", "m:alnScr", "m:degHide", "m:diff", "m:dispDef", "m:grow", "m:hideBot", "m:hideLeft",
    "m:hideRight", "m:hideTop", "m:lit", "m:maxDist", "m:noBreak", "m:nor", "m:objDist", "m:opEmu",
    "m:plcHide", "m:show", "m:smallFrac", "m:strikeBLTR", "m:strikeH", "m:strikeTLBR", "m:strikeV",
    "m:subHide", "m:supHide", "m:transp", "m:wrapRight", "m:zeroAsc", "m:zeroDesc", "m:zeroWid",
];
fn serialize(node: &Node, size: i64, language: Option<&str>) -> Result<String, String> {
    if let Some(text) = &node.text { return Ok(escape(text)); }
    if node.name == "w:rPr" { return Ok(run_properties(Some(node), size, language)); }
    if !node.name.starts_with("m:") { return Err(format!("Unsupported element in native math: {}", node.name)); }
    let mut xml = format!("<{}", node.name);
    for (k,v) in &node.attrs {
        // The shared writer also declares MathML's namespace; no element uses it.
        if k == "xmlns:w" || k == "xmlns:mml" { continue; }
        let value = if k == "m:val" && ON_OFF.contains(&node.name.as_str()) {
            match v.as_str() { "1" | "true" => "on", "0" | "false" => "off", _ => v.as_str() }
        } else { v.as_str() };
        xml.push_str(&format!(" {k}=\"{}\"", escape(value)));
    }
    xml.push('>');
    let mut inserted = node.children.iter().any(|n| n.name == "w:rPr");
    for child in &node.children {
        if node.name == "m:r" && !inserted && child.name != "m:rPr" && child.text.is_none() {
            xml.push_str(&run_properties(None, size, language)); inserted = true;
        }
        xml.push_str(&serialize(child, size, language)?);
    }
    xml.push_str(&format!("</{}>", node.name)); Ok(xml)
}
fn rels(kind: &str, target: &str) -> String {
    format!("<Relationships xmlns=\"{REL}\"><Relationship Id=\"rId1\" Type=\"{R}/{kind}\" Target=\"{target}\"/></Relationships>")
}
fn theme_elements() -> String {
    let colors = [("dk1","000000"),("lt1","FFFFFF"),("dk2","44546A"),("lt2","E7E6E6"),("accent1","4472C4"),("accent2","ED7D31"),("accent3","A5A5A5"),("accent4","FFC000"),("accent5","5B9BD5"),("accent6","70AD47"),("hlink","0563C1"),("folHlink","954F72")].iter().map(|(k,v)| format!("<a:{k}><a:srgbClr val=\"{v}\"/></a:{k}>")).collect::<String>();
    let fill = "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>";
    let fonts = "<a:latin typeface=\"Cambria Math\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/>";
    format!("<a:themeElements><a:clrScheme name=\"Office\">{colors}</a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont>{fonts}</a:majorFont><a:minorFont>{fonts}</a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst>{}</a:fillStyleLst><a:lnStyleLst>{}</a:lnStyleLst><a:effectStyleLst>{}</a:effectStyleLst><a:bgFillStyleLst>{}</a:bgFillStyleLst></a:fmtScheme></a:themeElements>", fill.repeat(3), format!("<a:ln w=\"12700\">{fill}<a:prstDash val=\"solid\"/></a:ln>").repeat(3), "<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3), fill.repeat(3))
}
fn theme() -> String {
    format!("<a:clipboardTheme xmlns:a=\"{A}\">{}<a:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/></a:clipboardTheme>", theme_elements())
}
/// The package's theme as a presentation theme part (WPS's equation slide).
fn presentation_theme() -> String {
    format!("<a:theme xmlns:a=\"{A}\" name=\"Office\">{}<a:objectDefaults/><a:extraClrSchemeLst/></a:theme>", theme_elements())
}

/// The equation's text paragraphs, as the clipboard package and WPS's equation
/// slide carry them: the a14:m paragraph (centered for a display equation) and
/// the empty paragraphs a display equation needs around it. Their a:, a14: and
/// m: prefixes are declared by the element that holds them. `language` is the
/// runs' language (a BCP 47 tag) from the text at the caret: next to WPS text,
/// runs without one got digits and operators tagged Arabic, which WPS draws
/// mirrored (VisualTeX for Windows' PowerPointMathClipboard.Entries).
pub(super) fn native_paragraphs(
    omml: &str, font_size: f64, display: bool, leading_paragraph: bool, trailing_paragraph: bool, language: Option<&str>,
) -> Result<String, String> {
    if !font_size.is_finite() || !(1.0..=512.0).contains(&font_size) { return Err("Invalid native equation font size".into()); }
    if !display && (leading_paragraph || trailing_paragraph) { return Err("Inline math must not introduce paragraphs".into()); }
    if language.is_some_and(|tag| tag.is_empty() || tag.len() > 35 || !tag.chars().all(|c| c.is_ascii_alphanumeric() || c == '-')) {
        return Err("Invalid native equation language".into());
    }
    let size = (font_size * 100.0).round() as i64;
    let math = serialize(&parse(omml)?,size,language)?;
    let mut paragraphs = String::new();
    if leading_paragraph { paragraphs.push_str("<a:p/>"); }
    paragraphs.push_str("<a:p>");
    if display { paragraphs.push_str("<a:pPr algn=\"ctr\"><a:buNone/></a:pPr>"); }
    paragraphs.push_str(&format!("<a14:m><m:oMathPara><m:oMathParaPr><m:jc m:val=\"centerGroup\"/></m:oMathParaPr>{math}</m:oMathPara></a14:m><a:endParaRPr sz=\"{size}\"/></a:p>"));
    if trailing_paragraph { paragraphs.push_str("<a:p/>"); }
    Ok(paragraphs)
}

/// An edit's paragraph for WPS: the edited equation's own zone (`zone`, its
/// a14:m element as WPS saved it) with the new equation in place of its own,
/// so it keeps its justification and inline or display form (VisualTeX for
/// Windows' PowerPointMathClipboard.EditEntries).
pub(super) fn native_edit_paragraph(omml: &str, font_size: f64, zone: &str, language: Option<&str>) -> Result<String, String> {
    let generated = parse_node(&format!(
        "<a:p xmlns:a=\"{A}\" xmlns:a14=\"{A14}\" xmlns:m=\"{M}\">{}</a:p>",
        native_paragraphs(omml, font_size, false, false, false, language)?
            .strip_prefix("<a:p>").and_then(|rest| rest.strip_suffix("</a:p>")).ok_or("Unexpected native paragraph")?
    ))?;
    let math = first_node(&generated, "m:oMath").ok_or("New native equation is missing")?.clone();
    let mut original = parse_node(zone).map_err(|_| "The edited native equation's zone is not valid XML".to_string())?;
    if original.name != "a14:m" { return Err("The edited equation is not inside a native math zone".into()); }
    fn count(node: &Node) -> usize {
        usize::from(node.name == "m:oMath") + node.children.iter().map(count).sum::<usize>()
    }
    if count(&original) != 1 { return Err("The edited native math zone does not hold exactly one equation".into()); }
    *first_node_mut(&mut original, "m:oMath").expect("counted above") = math;
    for (key, value) in [("xmlns:a", A), ("xmlns:a14", A14), ("xmlns:m", M)] {
        if !original.attrs.iter().any(|(k, _)| k == key) { original.attrs.push((key.into(), value.into())); }
    }
    let size = (font_size * 100.0).round() as i64;
    Ok(format!("<a:p>{}<a:endParaRPr sz=\"{size}\"/></a:p>", serialize_raw(&original)))
}

/// A one-slide presentation whose only text box holds `paragraphs`. WPS
/// Presentation pastes nothing from PowerPoint's equation clipboard, but
/// Slides.InsertFromFile imports a14:m equations from a presentation file as
/// native equations, and its own copy of that text pastes them anywhere
/// (VisualTeX for Windows' WpsEquationSlide).
pub(super) fn wps_equation_slide_entries(paragraphs: &str) -> Vec<(String, String)> {
    wps_equation_slide_boxes(&[paragraphs.to_string()])
}

/// The one-slide presentation with one text box per equation (a redraw writes
/// them all from one import): box n holds `boxes[n - 1]`.
pub(super) fn wps_equation_slide_boxes(boxes: &[String]) -> Vec<(String, String)> {
    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const DECLARATION: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    const COLOR_MAP: &str = "<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>";
    const EMPTY_TREE: &str = "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>";
    const SLIDE_TYPE: &str = "application/vnd.openxmlformats-officedocument.presentationml";
    let relationships = |items: &[(&str, &str, &str)]| {
        format!(
            "{DECLARATION}<Relationships xmlns=\"{REL}\">{}</Relationships>",
            items.iter().map(|(id, kind, target)| format!("<Relationship Id=\"{id}\" Type=\"{R}/{kind}\" Target=\"{target}\"/>")).collect::<String>()
        )
    };
    // The math shape is written the way PowerPoint saves one: under an a14
    // choice. The box only carries the text; WPS takes its own layout.
    let shapes: String = boxes
        .iter()
        .enumerate()
        .map(|(index, paragraphs)| {
            let id = index + 2;
            let y = 914400 * (1 + index % 5);
            format!(
                "<mc:AlternateContent xmlns:mc=\"{MC}\"><mc:Choice Requires=\"a14\"><p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"VisualTeX Equation {}\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>\
<p:spPr><a:xfrm><a:off x=\"914400\" y=\"{y}\"/><a:ext cx=\"9144000\" cy=\"914400\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>\
<p:txBody><a:bodyPr wrap=\"none\"><a:spAutoFit/></a:bodyPr><a:lstStyle/>{paragraphs}</p:txBody></p:sp></mc:Choice></mc:AlternateContent>",
                index + 1
            )
        })
        .collect();
    let slide = format!(
        "{DECLARATION}<p:sld xmlns:a=\"{A}\" xmlns:r=\"{R}\" xmlns:p=\"{P}\" xmlns:a14=\"{A14}\" xmlns:m=\"{M}\"><p:cSld><p:spTree>{EMPTY_TREE}{shapes}\
</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    );
    vec![
        ("[Content_Types].xml".into(), format!(
            "{DECLARATION}<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>\
<Override PartName=\"/ppt/presentation.xml\" ContentType=\"{SLIDE_TYPE}.presentation.main+xml\"/>\
<Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"{SLIDE_TYPE}.slide+xml\"/>\
<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"{SLIDE_TYPE}.slideLayout+xml\"/>\
<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"{SLIDE_TYPE}.slideMaster+xml\"/>\
<Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/></Types>")),
        ("_rels/.rels".into(), relationships(&[("rId1", "officeDocument", "ppt/presentation.xml")])),
        ("ppt/presentation.xml".into(), format!(
            "{DECLARATION}<p:presentation xmlns:a=\"{A}\" xmlns:r=\"{R}\" xmlns:p=\"{P}\"><p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id=\"256\" r:id=\"rId2\"/></p:sldIdLst><p:sldSz cx=\"12192000\" cy=\"6858000\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/></p:presentation>")),
        ("ppt/_rels/presentation.xml.rels".into(), relationships(&[
            ("rId1", "slideMaster", "slideMasters/slideMaster1.xml"),
            ("rId2", "slide", "slides/slide1.xml"),
            ("rId3", "theme", "theme/theme1.xml"),
        ])),
        ("ppt/slideMasters/slideMaster1.xml".into(), format!(
            "{DECLARATION}<p:sldMaster xmlns:a=\"{A}\" xmlns:r=\"{R}\" xmlns:p=\"{P}\"><p:cSld><p:spTree>{EMPTY_TREE}</p:spTree></p:cSld>{COLOR_MAP}<p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rId1\"/></p:sldLayoutIdLst></p:sldMaster>")),
        ("ppt/slideMasters/_rels/slideMaster1.xml.rels".into(), relationships(&[
            ("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"),
            ("rId2", "theme", "../theme/theme1.xml"),
        ])),
        ("ppt/slideLayouts/slideLayout1.xml".into(), format!(
            "{DECLARATION}<p:sldLayout xmlns:a=\"{A}\" xmlns:r=\"{R}\" xmlns:p=\"{P}\" type=\"blank\"><p:cSld name=\"Blank\"><p:spTree>{EMPTY_TREE}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>")),
        ("ppt/slideLayouts/_rels/slideLayout1.xml.rels".into(), relationships(&[("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")])),
        ("ppt/slides/slide1.xml".into(), slide),
        ("ppt/slides/_rels/slide1.xml.rels".into(), relationships(&[("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml")])),
        ("ppt/theme/theme1.xml".into(), format!("{DECLARATION}{}", presentation_theme())),
    ]
}

/// Native Art--Text-ClipFormat contains a DrawingML text fragment. Empty boundary
/// paragraphs represent explicit separators, not placeholder characters. There
/// is deliberately no plain-text/image fallback that could silently degrade math.
pub(super) fn clipboard_entries(
    omml: &str, font_size: f64, display: bool, leading_paragraph: bool, trailing_paragraph: bool,
) -> Result<Vec<(String, String)>, String> {
    let paragraphs = native_paragraphs(omml, font_size, display, leading_paragraph, trailing_paragraph, None)?;
    let xfrm = "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>";
    let drawing = format!("<a:graphic xmlns:a=\"{A}\" xmlns:a14=\"{A14}\" xmlns:m=\"{M}\"><a:graphicData uri=\"{LC}\"><lc:lockedCanvas xmlns:lc=\"{LC}\"><a:nvGrpSpPr><a:cNvPr id=\"0\" name=\"\"/><a:cNvGrpSpPr/></a:nvGrpSpPr><a:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></a:grpSpPr><a:sp><a:nvSpPr><a:cNvPr id=\"0\" name=\"\"/><a:cNvSpPr/></a:nvSpPr><a:spPr>{xfrm}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></a:spPr><a:txSp><a:txBody><a:bodyPr><a:noAutofit/></a:bodyPr><a:lstStyle/>{paragraphs}</a:txBody>{xfrm}</a:txSp></a:sp></lc:lockedCanvas></a:graphicData></a:graphic>");
    Ok(vec![
        ("[Content_Types].xml".into(), "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/clipboard/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/clipboard/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/></Types>".into()),
        ("_rels/.rels".into(), rels("drawing", "clipboard/drawings/drawing1.xml")),
        ("clipboard/drawings/drawing1.xml".into(), drawing),
        ("clipboard/drawings/_rels/drawing1.xml.rels".into(), rels("theme", "../theme/theme1.xml")),
        ("clipboard/theme/theme1.xml".into(), theme()),
    ])
}

fn first_node<'a>(node: &'a Node, name: &str) -> Option<&'a Node> {
    if node.name == name { return Some(node); }
    node.children.iter().find_map(|child| first_node(child, name))
}
fn first_node_mut<'a>(node: &'a mut Node, name: &str) -> Option<&'a mut Node> {
    if node.name == name { return Some(node); }
    node.children.iter_mut().find_map(|child| first_node_mut(child, name))
}
fn serialize_raw(node: &Node) -> String {
    if let Some(text) = &node.text { return escape(text); }
    let attrs=node.attrs.iter().map(|(k,v)|format!(" {k}=\"{}\"",escape(v))).collect::<String>();
    format!("<{}{attrs}>{}</{}>",node.name,node.children.iter().map(serialize_raw).collect::<String>(),node.name)
}

/// Preserve the original leading paragraph through Office's own complete native
/// text interchange. This avoids reimplementing bullets, tabs, indents and theme
/// inheritance property by property. All original paragraph TEXT is removed;
/// only its formatting and required native theme/media relationships are reused.
pub(super) fn preserve_leading_paragraph(
    generated: Vec<(String,String)>, original_clipboard: &[u8], paragraph_index: usize,
) -> Result<Vec<(String,Vec<u8>)>,String> {
    const MAX_BYTES:u64=32*1024*1024;
    const DRAWING:&str="clipboard/drawings/drawing1.xml";
    if original_clipboard.len() as u64 > MAX_BYTES { return Err("PowerPoint paragraph clipboard exceeds the size limit".into()); }
    let mut archive=zip::ZipArchive::new(Cursor::new(original_clipboard)).map_err(|e|format!("Invalid PowerPoint paragraph clipboard: {e}"))?;
    if archive.len()>128 { return Err("PowerPoint paragraph clipboard has too many parts".into()); }
    let mut entries=Vec::new(); let mut total=0_u64;
    for index in 0..archive.len() {
        let mut part=archive.by_index(index).map_err(|e|e.to_string())?;
        if part.is_dir() { continue; }
        if part.enclosed_name().is_none() || part.size()>MAX_BYTES || total.saturating_add(part.size())>MAX_BYTES {
            return Err("PowerPoint paragraph clipboard contains an invalid or oversized part".into());
        }
        let name=part.name().to_string();
        if entries.iter().any(|(n,_)|n==&name) { return Err("PowerPoint paragraph clipboard contains duplicate parts".into()); }
        let mut bytes=Vec::new(); part.by_ref().take(MAX_BYTES-total+1).read_to_end(&mut bytes).map_err(|e|e.to_string())?;
        total+=bytes.len() as u64;
        if total>MAX_BYTES { return Err("PowerPoint paragraph clipboard exceeds the expanded size limit".into()); }
        entries.push((name,bytes));
    }
    let slot=entries.iter().position(|(n,_)|n==DRAWING).ok_or("PowerPoint paragraph clipboard has no text drawing")?;
    let original=parse_node(std::str::from_utf8(&entries[slot].1).map_err(|_|"Invalid clipboard XML encoding")?)?;
    let generated_xml=&generated.iter().find(|(n,_)|n==DRAWING).ok_or("Native math clipboard is missing its drawing")?.1;
    let mut fresh=parse_node(generated_xml)?;
    let source_body=first_node(&original,"a:txBody").ok_or("Original shape clipboard contains no native text body")?;
    let source_paragraph=source_body.children.iter().filter(|n|n.name=="a:p").nth(paragraph_index.checked_sub(1).ok_or("Invalid paragraph index")?).ok_or("Original shape clipboard does not contain the target paragraph")?;
    let target_body=first_node_mut(&mut fresh,"a:txBody").ok_or("Native math text body is missing")?;
    let leading=target_body.children.iter_mut().find(|n|n.name=="a:p").ok_or("Missing leading paragraph")?;
    if !leading.children.is_empty() { return Err("Expected an empty leading paragraph before display math".into()); }
    leading.children=source_paragraph.children.iter().filter(|n|matches!(n.name.as_str(),"a:pPr"|"a:endParaRPr")).cloned().collect();
    if let Some(style)=source_body.children.iter().find(|n|n.name=="a:lstStyle") {
        if let Some(target_style)=target_body.children.iter_mut().find(|n|n.name=="a:lstStyle") { *target_style=style.clone(); }
    }
    fn namespaces(node:&Node, bindings:&mut Vec<(String,String)>) -> Result<(),String> {
        for (key,value) in node.attrs.iter().filter(|(k,_)|k.starts_with("xmlns:")) {
            if let Some((_,previous))=bindings.iter().find(|(k,_)|k==key) {
                if previous!=value { return Err("Ambiguous namespace in PowerPoint paragraph format".into()); }
            } else { bindings.push((key.clone(),value.clone())); }
        }
        for child in &node.children { namespaces(child,bindings)?; } Ok(())
    }
    let mut bindings=Vec::new(); namespaces(&original,&mut bindings)?;
    for (key,value) in bindings { if !fresh.attrs.iter().any(|(k,_)|k==&key) { fresh.attrs.push((key,value)); } }
    // The final drawing is always the generated TEXT envelope. The source
    // shape's geometry, nonvisual properties and text never enter it.
    entries[slot].1=serialize_raw(&fresh).into_bytes();
    Ok(entries)
}

/// Extract the actual mathematical contents from PowerPoint's native TEXT
/// clipboard. No name, alt text, tag, GUID, placeholder or formula registry is
/// consulted. Several adjacent oMath siblings in the one copied native zone are
/// serialized together, following Office's current natural merge boundary.
pub(super) fn native_equation_from_clipboard(bytes: &[u8]) -> Result<String, String> {
    const DRAWING: &str = "clipboard/drawings/drawing1.xml";
    const MAX: u64 = 4 * 1024 * 1024;
    if bytes.len() > 32 * 1024 * 1024 { return Err("Native equation clipboard is too large".into()); }
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).map_err(|e| format!("Invalid native equation clipboard: {e}"))?;
    let mut part = archive.by_name(DRAWING).map_err(|_| "The copied object has no native mathematical text")?;
    if part.size() > MAX { return Err("Native equation XML is too large".into()); }
    let mut xml = String::new();
    part.by_ref().take(MAX + 1).read_to_string(&mut xml).map_err(|e| e.to_string())?;
    if xml.len() as u64 > MAX { return Err("Native equation XML is too large".into()); }
    let root = parse_node(&xml)?;
    let body = first_node(&root,"a:txBody").ok_or("The copied object has no native text body")?;
    let mut equation = Node { name: "m:oMath".into(), attrs: vec![
        ("xmlns:m".into(), M.into()), ("xmlns:a".into(), A.into()),
        ("xmlns:w".into(), "http://schemas.openxmlformats.org/wordprocessingml/2006/main".into()),
    ], ..Node::default() };
    fn collect(node: &Node, output: &mut Vec<Node>) -> Result<(),String> {
        if node.name == "m:oMath" { output.extend(node.children.clone()); return Ok(()); }
        if node.name == "a:r" {
            let text=node.children.iter().filter(|n|n.name=="a:t").flat_map(|n|n.children.iter().filter_map(|n|n.text.as_deref())).collect::<String>();
            if text.chars().all(|c|matches!(c,' '|'\t'|'\u{a0}')) {return Ok(());}
            return Err("The copied range contains ordinary text outside the selected equation".into());
        }
        if node.name == "a:fld" || node.name == "a:br" {
            return Err("The copied range contains ordinary text outside the selected equation".into());
        }
        if matches!(node.name.as_str(),"a:pPr"|"a:endParaRPr"|"a:bodyPr"|"a:lstStyle"|"m:oMathParaPr") { return Ok(()); }
        for child in &node.children { collect(child,output)?; }
        Ok(())
    }
    let mut nonempty_paragraphs = 0;
    for paragraph in body.children.iter().filter(|n| n.name == "a:p") {
        let before = equation.children.len();
        collect(paragraph,&mut equation.children)?;
        if equation.children.len() > before { nonempty_paragraphs += 1; }
    }
    if nonempty_paragraphs != 1 { return Err("Select exactly one current native equation, not several paragraphs".into()); }
    // Copy-generated volatile proofing/layout flags are not mathematical edits.
    fn canonicalize(node:&mut Node) {
        node.attrs.retain(|(k,_)| !matches!(k.as_str(),"dirty"|"smtClean"));
        node.attrs.sort();
        for child in &mut node.children { canonicalize(child); }
    }
    canonicalize(&mut equation);
    Ok(serialize_raw(&equation))
}

/// Replace only the native math in the exact original text clipboard fragment.
/// Visible whitespace runs keep their original fonts/styles; the clipboard
/// envelope is never added as a shape. The application can still merge native
/// math freely when there is no intervening ordinary text.
pub(super) fn replace_native_edit_context(
    generated: Vec<(String,String)>, original_clipboard:&[u8],
) -> Result<Vec<(String,Vec<u8>)>,String> {
    const DRAWING:&str="clipboard/drawings/drawing1.xml";
    const MAX:u64=32*1024*1024;
    // Validate that this fragment contains exactly one native mathematical zone
    // with, at most, its visible surrounding whitespace.
    native_equation_from_clipboard(original_clipboard)?;
    let generated_root=parse_node(&generated.iter().find(|(n,_)|n==DRAWING).ok_or("New native clipboard has no drawing")?.1)?;
    let math=first_node(&generated_root,"m:oMath").ok_or("New native equation is missing")?.clone();
    let mut archive=zip::ZipArchive::new(Cursor::new(original_clipboard)).map_err(|e|e.to_string())?;
    if archive.len()>128 {return Err("The native clipboard has too many parts".into());}
    let mut entries=Vec::new(); let mut total=0_u64;
    for index in 0..archive.len() {
        let mut part=archive.by_index(index).map_err(|e|e.to_string())?;
        if part.is_dir(){continue;}
        if part.enclosed_name().is_none()||part.size()>MAX||total.saturating_add(part.size())>MAX {return Err("Invalid native clipboard part".into());}
        let name=part.name().to_string();let mut data=Vec::new();
        part.by_ref().take(MAX-total+1).read_to_end(&mut data).map_err(|e|e.to_string())?;
        total+=data.len() as u64;
        if total>MAX {return Err("Native clipboard exceeds expanded size limit".into());}
        if entries.iter().any(|(n,_)|n==&name) {return Err("Duplicate native clipboard part".into());}
        entries.push((name,data));
    }
    let slot=entries.iter().position(|(n,_)|n==DRAWING).ok_or("Original native drawing is missing")?;
    let mut root=parse_node(std::str::from_utf8(&entries[slot].1).map_err(|_|"Invalid native XML")?)?;
    fn replace(node:&mut Node, math:&Node, inserted:&mut bool) {
        let mut children=Vec::new();
        for mut child in std::mem::take(&mut node.children) {
            if child.name=="m:oMath" {
                if !*inserted {children.push(math.clone()); *inserted=true;}
            } else {replace(&mut child,math,inserted); children.push(child);}
        }
        node.children=children;
    }
    let mut inserted=false; replace(&mut root,&math,&mut inserted);
    if !inserted {return Err("Original math was not found in its native text fragment".into());}
    for (key,value) in [("xmlns:a",A),("xmlns:m",M),("xmlns:a14",A14)] {
        if !root.attrs.iter().any(|(k,_)|k==key) {root.attrs.push((key.into(),value.into()));}
    }
    entries[slot].1=serialize_raw(&root).into_bytes();
    Ok(entries)
}

#[cfg(test)]
mod tests {
    use super::*;
    fn sample() -> String { format!("<m:oMath xmlns:m=\"{M}\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><m:f><m:fPr/><m:num><m:r><m:t>x+1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>") }
    #[test] fn radicals_use_powerpoint_native_degree_visibility() {
        for (input, expected) in [("1", "on"), ("true", "on"), ("0", "off"), ("false", "off")] {
            let xml = format!("<m:oMath xmlns:m=\"{M}\"><m:rad><m:radPr><m:degHide m:val=\"{input}\"/></m:radPr><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad></m:oMath>");
            let entries = clipboard_entries(&xml, 24.0, false, false, false).unwrap();
            let drawing = &entries.iter().find(|(name, _)| name == "clipboard/drawings/drawing1.xml").unwrap().1;
            assert!(drawing.contains(&format!("<m:degHide m:val=\"{expected}\">")));
            assert!(drawing.contains("<m:t>3</m:t>"));
        }
    }
    // PowerPoint equations come from the shared writer (office::omml_write):
    // every Windows reference equation must package and read back.
    #[test] fn shared_writer_equations_package_for_powerpoint() {
        let rows: Vec<serde_json::Value> =
            serde_json::from_str(include_str!("omml_write/fixtures/windows_forward.json")).unwrap();
        assert!(rows.len() > 200);
        for row in rows {
            let omml = row["omml"].as_str().unwrap();
            let display = row["mathml"].as_str().unwrap().contains("display=\"block\"");
            let entries = clipboard_entries(omml, 24.0, display, false, false)
                .unwrap_or_else(|e| panic!("{}: {e}", row["source"]));
            let drawing = &entries.iter().find(|(name, _)| name == "clipboard/drawings/drawing1.xml").unwrap().1;
            assert!(!drawing.contains("xmlns:mml") && !drawing.contains("w:rPr"), "{}", row["source"]);
            let bytes = super::super::macos_offline::build_stored_zip(&entries).unwrap();
            native_equation_from_clipboard(&bytes).unwrap_or_else(|e| panic!("{}: {e}", row["source"]));
        }
    }
    // VisualTeX for Windows packs the same OMML into the same bytes
    // (VisualTeX.WindowsOffice.Tests PowerPointMathClipboardTests.PackageMatchesTheMacPackage).
    // Change both or neither.
    #[test] fn package_matches_the_windows_package() {
        use sha2::{Digest, Sha256};
        let xml = format!("<m:oMath xmlns:m=\"{M}\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><m:rad><m:radPr><m:degHide m:val=\"1\"/></m:radPr><m:deg/><m:e><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/><w:color w:val=\"FF0000\"/></w:rPr><m:t>a&lt;b&amp;c</m:t></m:r></m:e></m:rad></m:oMath>");
        let package = super::super::macos_offline::build_stored_zip(&clipboard_entries(&xml, 24.0, true, false, false).unwrap()).unwrap();
        assert_eq!(format!("{:x}", Sha256::digest(&package)), "abc0b2da908c1622ccc8c236064955abe2647ddb56c881801837d252fb66f1ec");
    }
    #[test] fn on_off_properties_use_powerpoint_canonical_values() {
        let xml = format!("<m:oMath xmlns:m=\"{M}\"><m:nary><m:naryPr><m:chr m:val=\"∮\"/><m:limLoc m:val=\"subSup\"/><m:grow m:val=\"1\"/><m:supHide m:val=\"1\"/><m:subHide m:val=\"0\"/></m:naryPr><m:sub><m:r><m:t>C</m:t></m:r></m:sub><m:sup/><m:e><m:r><m:t>F</m:t></m:r></m:e></m:nary><m:m><m:mPr><m:plcHide m:val=\"true\"/><m:mcs><m:mc><m:mcPr><m:count m:val=\"1\"/></m:mcPr></m:mc></m:mcs></m:mPr><m:mr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:mr></m:m></m:oMath>");
        let entries = clipboard_entries(&xml, 24.0, false, false, false).unwrap();
        let drawing = &entries.iter().find(|(name, _)| name == "clipboard/drawings/drawing1.xml").unwrap().1;
        for expected in ["<m:supHide m:val=\"on\">", "<m:subHide m:val=\"off\">", "<m:grow m:val=\"on\">", "<m:plcHide m:val=\"on\">"] {
            assert!(drawing.contains(expected), "{expected}");
        }
        // Other numeric and named values are not on/off properties.
        assert!(drawing.contains("<m:count m:val=\"1\">") && drawing.contains("<m:limLoc m:val=\"subSup\">"));
    }
    #[test] fn native_text_not_a_presentation_or_picture() {
        let e=clipboard_entries(&sample(),24.0,false,false,false).unwrap();
        assert_eq!(e.len(),5);
        let xml=&e.iter().find(|(n,_)|n=="clipboard/drawings/drawing1.xml").unwrap().1;
        assert_eq!(xml.matches("<m:oMath ").count(),1);
        assert_eq!(xml.matches("<a:rPr sz=\"2400\"").count(),2);
        assert!(!xml.contains("w:rPr")); assert!(!xml.contains("VisualTeX"));
        assert!(!xml.contains("<a:blip")); assert!(!xml.contains("<p:"));
    }
    #[test] fn reads_untagged_native_math_and_rejects_plain_text() {
        let entries=clipboard_entries(&sample(),24.0,false,false,false).unwrap();
        let bytes=super::super::macos_offline::build_stored_zip(&entries).unwrap();
        let equation=native_equation_from_clipboard(&bytes).unwrap();
        assert_eq!(equation.matches("<m:f>").count(),1);
        assert_eq!(to_latex::from_omml(&equation).unwrap().latex,r"\frac{x+1}{2}");
        assert!(!equation.contains("lc:") && !equation.contains("VisualTeX"));
        let mut invalid=entries.clone();
        invalid[2].1=invalid[2].1.replace("<a:p>","<a:p><a:r><a:t>text</a:t></a:r>");
        let invalid=super::super::macos_offline::build_stored_zip(&invalid).unwrap();
        assert!(native_equation_from_clipboard(&invalid).is_err());
        let mut merged=entries;
        merged[2].1=merged[2].1.replace("</m:oMathPara>",&format!("{}</m:oMathPara>",sample()));
        let merged=super::super::macos_offline::build_stored_zip(&merged).unwrap();
        assert_eq!(native_equation_from_clipboard(&merged).unwrap().matches("<m:f>").count(),2);
    }
    #[test] fn explicit_paragraph_boundaries_only_for_display() {
        for (l,r,n) in [(false,false,0),(true,false,1),(false,true,1),(true,true,2)] {
            let e=clipboard_entries(&sample(),24.0,true,l,r).unwrap();
            assert_eq!(e[2].1.matches("<a:p/>").count(),n);
            assert!(e[2].1.contains("algn=\"ctr\""));
        }
        assert!(clipboard_entries(&sample(),24.0,false,true,false).is_err());
    }
    #[test] fn rejects_non_math_and_invalid_sizes() {
        for input in ["<x/>","", "<!DOCTYPE x><m:oMath/>","<m:oMath/>"] { assert!(clipboard_entries(input,24.0,false,false,false).is_err()); }
        for size in [0.0,f64::NAN,513.0] { assert!(clipboard_entries(&sample(),size,false,false,false).is_err()); }
    }
    #[test] fn preserves_math_structure_fonts_and_escaped_symbols() {
        let source=sample().replace("<m:t>x+1</m:t>","<w:rPr><w:rFonts w:ascii=\"Times New Roman\"/><w:b/><w:color w:val=\"123ABC\"/></w:rPr><m:t>α &amp; β &amp;amp;</m:t>");
        let xml=serialize(&parse(&source).unwrap(),1800,None).unwrap();
        assert!(xml.contains("typeface=\"Times New Roman\"")); assert!(xml.contains("b=\"1\""));
        assert!(xml.contains("α &amp; β &amp;amp;")); assert!(xml.contains("<m:f>"));
    }
    #[test] fn preserves_full_paragraph_format_without_copying_original_text() {
        let mut source=clipboard_entries(&sample(),24.0,false,false,false).unwrap();
        source[2].1=source[2].1.replace("<a:p>","<a:p><a:pPr algn=\"r\"><a:spcBef><a:spcPts val=\"1100\"/></a:spcBef><a:buAutoNum type=\"arabicPeriod\"/></a:pPr><a:r><a:t>PRIVATE_ORIGINAL_TEXT</a:t></a:r>");
        let zip=super::super::macos_offline::build_stored_zip(&source).unwrap();
        let fresh=clipboard_entries(&sample(),24.0,true,true,true).unwrap();
        let result=preserve_leading_paragraph(fresh,&zip,1).unwrap();
        let xml=std::str::from_utf8(&result.iter().find(|(n,_)|n=="clipboard/drawings/drawing1.xml").unwrap().1).unwrap();
        assert!(!xml.contains("PRIVATE_ORIGINAL_TEXT")); assert_eq!(xml.matches("<m:oMath ").count(),1);
        assert!(xml.contains("algn=\"r\"")); assert!(xml.contains("val=\"1100\"")); assert!(xml.contains("arabicPeriod"));
        assert!(preserve_leading_paragraph(clipboard_entries(&sample(),24.0,true,true,true).unwrap(),b"invalid",1).is_err());
    }
    #[test] fn wps_equation_slide_is_a_one_slide_presentation_with_the_paragraphs() {
        let paragraphs = native_paragraphs(&sample(), 20.0, true, false, false, None).unwrap();
        assert!(paragraphs.starts_with("<a:p><a:pPr algn=\"ctr\">"));
        let entries = wps_equation_slide_entries(&paragraphs);
        let bytes = super::super::macos_offline::build_stored_zip(&entries).unwrap();
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        for index in 0..archive.len() {
            let mut part = archive.by_index(index).unwrap();
            let mut xml = String::new();
            part.read_to_string(&mut xml).unwrap();
            // Every part is well-formed XML.
            let mut reader = Reader::from_str(&xml);
            loop {
                match reader.read_event() {
                    Ok(Event::Eof) => break,
                    Ok(_) => {}
                    Err(error) => panic!("{}: {error}", part.name()),
                }
            }
            if part.name() == "ppt/slides/slide1.xml" {
                assert_eq!(xml.matches("<p:sp>").count(), 1);
                assert!(xml.contains("name=\"VisualTeX Equation 1\""));
                assert!(xml.contains("<mc:Choice Requires=\"a14\">") && xml.contains("<m:oMathPara>"));
                assert!(xml.contains("algn=\"ctr\"") && xml.contains("<m:t>x+1</m:t>"));
            }
            if part.name() == "ppt/theme/theme1.xml" {
                assert!(xml.contains("<a:theme ") && xml.contains("<a:objectDefaults/>"));
            }
        }
        assert_eq!(archive.len(), 11);
    }
    #[test] fn a_redraw_slide_has_one_box_per_equation() {
        let first = native_paragraphs(&sample(), 20.0, false, false, false, None).unwrap();
        let second = native_paragraphs(&sample(), 28.0, true, false, false, None).unwrap();
        let entries = wps_equation_slide_boxes(&[first, second]);
        let slide = &entries.iter().find(|(name, _)| name == "ppt/slides/slide1.xml").unwrap().1;
        assert_eq!(slide.matches("<p:sp>").count(), 2);
        assert!(slide.contains("<p:cNvPr id=\"2\" name=\"VisualTeX Equation 1\"/>"));
        assert!(slide.contains("<p:cNvPr id=\"3\" name=\"VisualTeX Equation 2\"/>"));
        assert!(slide.find("sz=\"2000\"").unwrap() < slide.find("sz=\"2800\"").unwrap());
    }
    #[test] fn runs_carry_the_text_language_when_given() {
        let omml = format!("<m:oMath xmlns:m=\"{M}\"><m:r><m:t>a</m:t></m:r><m:r><m:t>&gt;</m:t></m:r></m:oMath>");
        let tagged = native_paragraphs(&omml, 18.0, false, false, false, Some("en-US")).unwrap();
        assert_eq!(tagged.matches("<a:rPr lang=\"en-US\" altLang=\"en-US\" sz=\"1800\">").count(), 2);
        assert!(!native_paragraphs(&omml, 18.0, false, false, false, None).unwrap().contains("lang="));
        assert!(native_paragraphs(&omml, 18.0, false, false, false, Some("en\"US")).is_err());
    }
    #[test] fn an_edit_keeps_its_zone_and_replaces_only_the_equation() {
        let zone = format!("<a14:m xmlns:a14=\"{A14}\" xmlns:m=\"{M}\" xmlns:a=\"{A}\"><m:oMathPara><m:oMathParaPr><m:jc m:val=\"left\"/></m:oMathParaPr><m:oMath><m:r><m:t>old</m:t></m:r></m:oMath></m:oMathPara></a14:m>");
        let paragraph = native_edit_paragraph(&sample(), 18.0, &zone, None).unwrap();
        assert!(paragraph.starts_with("<a:p><a14:m "), "{paragraph}");
        assert!(paragraph.ends_with("<a:endParaRPr sz=\"1800\"/></a:p>"), "{paragraph}");
        assert!(paragraph.contains("<m:jc m:val=\"left\">") && paragraph.contains("<m:t>x+1</m:t>") && !paragraph.contains("old"));
        assert!(!paragraph.contains("algn="));
        let two = zone.replace("</m:oMathPara>", "<m:oMath><m:r><m:t>b</m:t></m:r></m:oMath></m:oMathPara>");
        assert!(native_edit_paragraph(&sample(), 18.0, &two, None).is_err());
        assert!(native_edit_paragraph(&sample(), 18.0, &format!("<m:oMath xmlns:m=\"{M}\"/>"), None).is_err());
        assert!(native_edit_paragraph(&sample(), 18.0, "<a14:m", None).is_err());
    }
    #[test] fn style_owned_ui_fixture_when_requested() {
        if let Ok(input)=std::env::var("VISUALTEX_PPT_CONTEXT_FIXTURE_IN") {
            let output=std::env::var("VISUALTEX_PPT_CONTEXT_FIXTURE_OUT").unwrap();
            let trailing=std::env::var("VISUALTEX_PPT_CONTEXT_TRAILING").unwrap_or_default()=="1";
            let context=std::fs::read(input).unwrap();
            let fresh=clipboard_entries(&sample(),24.0,true,true,trailing).unwrap();
            let index=std::env::var("VISUALTEX_PPT_CONTEXT_PARAGRAPH").unwrap_or("1".into()).parse().unwrap();
            let entries=preserve_leading_paragraph(fresh,&context,index).unwrap();
            std::fs::write(output,super::super::macos_offline::build_stored_zip(&entries).unwrap()).unwrap();
        }
    }
    #[test]
    #[ignore = "builds native clipboard fixtures from the actual frontend renderer"]
    fn build_native_edit_corpus() {
        use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine as _};
        let source=std::env::var("VISUALTEX_NATIVE_EDIT_CORPUS").unwrap();
        let samples:Vec<serde_json::Value>=serde_json::from_slice(&std::fs::read(&source).unwrap()).unwrap();
        let directory=std::path::Path::new(&source).parent().unwrap();
        for sample in samples {
            let name=sample["name"].as_str().unwrap();
            assert!(name.chars().all(|c|c.is_ascii_alphanumeric()||c=='_'));
            let xml=String::from_utf8(URL_SAFE_NO_PAD.decode(sample["ommlBase64"].as_str().unwrap()).unwrap()).unwrap();
            let entries=clipboard_entries(&xml,24.0,false,false,false).unwrap();
            std::fs::write(directory.join(format!("{name}.bin")),super::super::macos_offline::build_stored_zip(&entries).unwrap()).unwrap();
        }
    }

    #[test] fn write_owned_ui_fixtures_when_requested() {
        if let Ok(directory)=std::env::var("VISUALTEX_PPT_NATIVE_FIXTURE_DIR") {
            for (name,size,display,l,r) in [("inline",24.0,false,false,false),("inline36",36.0,false,false,false),("display00",24.0,true,false,false),("display10",24.0,true,true,false),("display01",24.0,true,false,true),("display11",24.0,true,true,true)] {
                let entries=clipboard_entries(&sample(),size,display,l,r).unwrap();
                let bytes=super::super::macos_offline::build_stored_zip(&entries).unwrap();
                std::fs::write(std::path::Path::new(&directory).join(format!("native-{name}.bin")),bytes).unwrap();
            }
        }
    }
}
