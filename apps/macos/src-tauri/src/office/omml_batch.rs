use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine as _};
use quick_xml::{events::Event, Reader};
use std::{
    collections::HashMap,
    fs,
    path::{Path, PathBuf},
};

use super::macos_offline::word_omml_to_mathml;

const CLI_ARGUMENT: &str = "--office-omml-to-latex-batch";
const PROTOCOL_VERSION: &str = "1";
const MAX_BATCH_BYTES: usize = 16 * 1024 * 1024;
const MAX_ITEMS: usize = 1000;

#[derive(Debug, Default, Clone)]
struct XmlNode {
    name: String,
    attributes: HashMap<String, String>,
    text: String,
    children: Vec<XmlNode>,
}

impl XmlNode {
    fn attribute(&self, name: &str) -> Option<&str> {
        self.attributes.get(name).map(String::as_str)
    }

    fn all_text(&self) -> String {
        let mut result = self.text.clone();
        for child in &self.children {
            result.push_str(&child.all_text());
        }
        result
    }
}

pub fn run_cli_if_requested() -> Option<i32> {
    let arguments = std::env::args_os().collect::<Vec<_>>();
    let position = arguments
        .iter()
        .position(|value| value == std::ffi::OsStr::new(CLI_ARGUMENT))?;
    let Some(input) = arguments.get(position + 1).map(PathBuf::from) else {
        eprintln!("{CLI_ARGUMENT} requires an input path");
        return Some(2);
    };
    let Some(output) = arguments.get(position + 2).map(PathBuf::from) else {
        eprintln!("{CLI_ARGUMENT} requires an output path");
        return Some(2);
    };
    let status = match convert_manifest(&input, &output) {
        Ok(()) => 0,
        Err(error) => {
            eprintln!("OMML batch conversion failed: {error}");
            2
        }
    };
    Some(status)
}

fn parse_manifest(path: &Path) -> Result<HashMap<String, String>, String> {
    let metadata = fs::metadata(path)
        .map_err(|error| format!("Unable to inspect {}: {error}", path.display()))?;
    if !metadata.is_file() || metadata.len() == 0 || metadata.len() as usize > MAX_BATCH_BYTES {
        return Err("The OMML batch manifest has an invalid size".to_string());
    }
    let contents = fs::read_to_string(path)
        .map_err(|error| format!("Unable to read {}: {error}", path.display()))?;
    let mut values = HashMap::new();
    for row in contents.lines().filter(|row| !row.is_empty()) {
        let (key, value) = row
            .split_once('=')
            .ok_or_else(|| "The OMML batch manifest contains an invalid row".to_string())?;
        if key.is_empty()
            || !key.bytes().all(|value| value.is_ascii_alphanumeric())
            || values.insert(key.to_string(), value.to_string()).is_some()
        {
            return Err("The OMML batch manifest contains an invalid key".to_string());
        }
    }
    Ok(values)
}

fn required<'a>(values: &'a HashMap<String, String>, key: &str) -> Result<&'a str, String> {
    values
        .get(key)
        .map(String::as_str)
        .ok_or_else(|| format!("The OMML batch manifest is missing {key}"))
}

fn decode_utf8(value: &str, label: &str) -> Result<String, String> {
    let bytes = URL_SAFE_NO_PAD
        .decode(value)
        .map_err(|error| format!("Unable to decode {label}: {error}"))?;
    String::from_utf8(bytes).map_err(|_| format!("{label} is not valid UTF-8"))
}

fn convert_manifest(input: &Path, output: &Path) -> Result<(), String> {
    let values = parse_manifest(input)?;
    if required(&values, "protocolVersion")? != PROTOCOL_VERSION {
        return Err("The OMML batch protocol version is invalid".to_string());
    }
    if required(&values, "sourceKind")? != "omml" {
        return Err("The OMML batch source kind is invalid".to_string());
    }
    let item_count = required(&values, "itemCount")?
        .parse::<usize>()
        .map_err(|_| "The OMML batch item count is invalid".to_string())?;
    if item_count == 0 || item_count > MAX_ITEMS {
        return Err(format!("The OMML batch must contain 1 to {MAX_ITEMS} formulas"));
    }

    let session_id = required(&values, "sessionId")?;
    let bookmark_name = required(&values, "bookmarkName")?;
    let source_document_id = decode_utf8(
        required(&values, "sourceDocumentId")?,
        "the Word document identity",
    )?;

    let mut output_rows = vec![
        format!("protocolVersion={PROTOCOL_VERSION}"),
        format!("sessionId={session_id}"),
        "operation=formulaRestore".to_string(),
        "outputKind=latex".to_string(),
        format!("sourceDocumentId={source_document_id}"),
        format!("bookmarkName={bookmark_name}"),
        format!("itemCount={item_count}"),
    ];

    for index in 0..item_count {
        let prefix = format!("item{index}");
        let display_mode = required(&values, &format!("{prefix}displayMode"))?;
        if !matches!(display_mode, "inline" | "block") {
            return Err(format!("Formula {index} has an invalid display mode"));
        }
        let omml = decode_utf8(
            required(&values, &format!("{prefix}payloadBase64"))?,
            &format!("formula {index} OMML"),
        )?;
        let math_ml = word_omml_to_mathml(&omml)?;
        let latex = mathml_to_latex(&math_ml)?;
        let wrapped = if display_mode == "block" {
            format!("$${latex}$$")
        } else {
            format!("${latex}$")
        };
        output_rows.extend([
            format!("{prefix}kind=text"),
            format!("{prefix}textBase64={}", URL_SAFE_NO_PAD.encode(wrapped.as_bytes())),
            format!(
                "{prefix}sourceStart={}",
                required(&values, &format!("{prefix}sourceStart"))?
            ),
            format!(
                "{prefix}sourceEnd={}",
                required(&values, &format!("{prefix}sourceEnd"))?
            ),
            format!(
                "{prefix}sourceTextBase64={}",
                required(&values, &format!("{prefix}sourceTextBase64"))?
            ),
        ]);
    }

    write_output_manifest(output, &output_rows)
}

fn write_output_manifest(output: &Path, output_rows: &[String]) -> Result<(), String> {
    let result = format!("{}\n", output_rows.join("\n"));
    if result.len() > MAX_BATCH_BYTES {
        return Err("The OMML batch result exceeds 16 MB".to_string());
    }
    let parent = output
        .parent()
        .ok_or_else(|| "The OMML batch output path has no parent".to_string())?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
    let temporary = output.with_extension("tmp");
    fs::write(&temporary, result)
        .map_err(|error| format!("Unable to write {}: {error}", temporary.display()))?;
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        fs::set_permissions(&temporary, fs::Permissions::from_mode(0o600))
            .map_err(|error| format!("Unable to protect {}: {error}", temporary.display()))?;
    }
    fs::rename(&temporary, output)
        .map_err(|error| format!("Unable to publish {}: {error}", output.display()))?;
    Ok(())
}

fn local_name(value: &[u8]) -> String {
    let value = String::from_utf8_lossy(value);
    value.rsplit(':').next().unwrap_or(&value).to_string()
}

fn parse_xml(value: &str) -> Result<XmlNode, String> {
    if value.is_empty()
        || value.len() > 4_000_000
        || value.contains("<!DOCTYPE")
        || value.contains("<!ENTITY")
    {
        return Err("Word returned invalid or excessive MathML".to_string());
    }
    let mut reader = Reader::from_str(value);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::<XmlNode>::new();
    let mut root = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                let mut node = XmlNode {
                    name: local_name(start.name().as_ref()),
                    ..XmlNode::default()
                };
                for attribute in start.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| error.to_string())?;
                    let key = local_name(attribute.key.as_ref());
                    let value = attribute
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(|error| error.to_string())?
                        .into_owned();
                    node.attributes.insert(key, value);
                }
                stack.push(node);
            }
            Ok(Event::Empty(start)) => {
                let mut node = XmlNode {
                    name: local_name(start.name().as_ref()),
                    ..XmlNode::default()
                };
                for attribute in start.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| error.to_string())?;
                    let key = local_name(attribute.key.as_ref());
                    let value = attribute
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(|error| error.to_string())?
                        .into_owned();
                    node.attributes.insert(key, value);
                }
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(node);
                } else if root.replace(node).is_some() {
                    return Err("MathML contains multiple root elements".to_string());
                }
            }
            Ok(Event::Text(text)) => {
                if let Some(parent) = stack.last_mut() {
                    let decoded = text
                        .decode()
                        .map_err(|error| error.to_string())?
                        .into_owned();
                    parent.text.push_str(&decoded);
                }
            }
            Ok(Event::GeneralRef(reference)) => {
                if let Some(parent) = stack.last_mut() {
                    let value = format!("&{};", reference.decode().map_err(|e|e.to_string())?);
                    parent.text.push_str(&quick_xml::escape::unescape(&value).map_err(|e|e.to_string())?);
                }
            }
            Ok(Event::CData(text)) => {
                if let Some(parent) = stack.last_mut() {
                    parent.text.push_str(
                        &text
                            .decode()
                            .map_err(|error| error.to_string())?
                            .into_owned(),
                    );
                }
            }
            Ok(Event::End(_)) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| "MathML contains an unexpected closing tag".to_string())?;
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(node);
                } else if root.replace(node).is_some() {
                    return Err("MathML contains multiple root elements".to_string());
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(format!("Unable to parse Word MathML: {error}")),
        }
    }
    if !stack.is_empty() {
        return Err("Word MathML is incomplete".to_string());
    }
    root.ok_or_else(|| "Word MathML has no root element".to_string())
}

fn escape_latex(value: &str, text_mode: bool) -> String {
    let mut result = String::new();
    for character in value.chars() {
        match character {
            '\\' if text_mode => result.push_str("\\textbackslash{}"),
            '\\' => result.push_str("\\backslash "),
            '^' if text_mode => result.push_str("\\textasciicircum{}"),
            '~' if text_mode => result.push_str("\\textasciitilde{}"),
            '$' | '{' | '}' | '#' | '%' | '&' | '_' => {
                result.push('\\');
                result.push(character);
            }
            _ => result.push(character),
        }
    }
    result
}

fn token(value: &str) -> String {
    let value = value.trim();
    let mapped = match value {
        "\u{2061}" | "\u{2062}" | "\u{2064}" => "", "\u{2063}" => ",",
        "−" => "-", "×" => "\\times ", "÷" => "\\div ", "±" => "\\pm ",
        "∓" => "\\mp ", "·" => "\\cdot ", "∗" => "\\ast ", "∘" => "\\circ ",
        "∞" => "\\infty ", "∂" => "\\partial ", "∇" => "\\nabla ",
        "∑" => "\\sum ", "∏" => "\\prod ", "∫" => "\\int ",
        "∬" => "\\iint ", "∭" => "\\iiint ", "∮" => "\\oint ",
        "√" => "\\sqrt{}", "≈" => "\\approx ", "≃" => "\\simeq ",
        "≅" => "\\cong ", "≠" => "\\ne ", "≤" => "\\le ", "≥" => "\\ge ",
        "≪" => "\\ll ", "≫" => "\\gg ", "≡" => "\\equiv ", "∝" => "\\propto ",
        "∈" => "\\in ", "∉" => "\\notin ", "∋" => "\\ni ",
        "⊂" => "\\subset ", "⊃" => "\\supset ", "⊆" => "\\subseteq ",
        "⊇" => "\\supseteq ", "∪" => "\\cup ", "∩" => "\\cap ",
        "∧" => "\\land ", "∨" => "\\lor ", "¬" => "\\neg ",
        "∀" => "\\forall ", "∃" => "\\exists ", "∄" => "\\nexists ",
        "∅" => "\\varnothing ", "⊥" => "\\perp ", "∥" => "\\parallel ",
        "←" => "\\leftarrow ", "→" => "\\rightarrow ", "↔" => "\\leftrightarrow ",
        "⇐" => "\\Leftarrow ", "⇒" => "\\Rightarrow ", "⇔" => "\\Leftrightarrow ",
        "↦" => "\\mapsto ", "α" => "\\alpha ", "β" => "\\beta ",
        "γ" => "\\gamma ", "δ" => "\\delta ", "ε" => "\\epsilon ",
        "ϵ" => "\\varepsilon ", "ζ" => "\\zeta ", "η" => "\\eta ",
        "θ" => "\\theta ", "ϑ" => "\\vartheta ", "ι" => "\\iota ",
        "κ" => "\\kappa ", "λ" => "\\lambda ", "μ" => "\\mu ",
        "ν" => "\\nu ", "ξ" => "\\xi ", "π" => "\\pi ",
        "ϖ" => "\\varpi ", "ρ" => "\\rho ", "ϱ" => "\\varrho ",
        "σ" => "\\sigma ", "ς" => "\\varsigma ", "τ" => "\\tau ",
        "υ" => "\\upsilon ", "φ" => "\\phi ", "ϕ" => "\\varphi ",
        "χ" => "\\chi ", "ψ" => "\\psi ", "ω" => "\\omega ",
        "Γ" => "\\Gamma ", "Δ" => "\\Delta ", "Θ" => "\\Theta ",
        "Λ" => "\\Lambda ", "Ξ" => "\\Xi ", "Π" => "\\Pi ",
        "Σ" => "\\Sigma ", "Υ" => "\\Upsilon ", "Φ" => "\\Phi ",
        "Ψ" => "\\Psi ", "Ω" => "\\Omega ", _ => return escape_latex(value, false),
    };
    mapped.to_string()
}

pub(super) fn native_math_token(value: &str) -> String { token(value) }

fn children(node: &XmlNode) -> String {
    node.children.iter().map(convert_node).collect::<String>()
}

fn group(value: String) -> String {
    let value = value.trim().to_string();
    if value.chars().count() == 1 || (value.starts_with('\\') && value[1..].chars().all(|c| c.is_ascii_alphabetic())) {
        value
    } else {
        format!("{{{value}}}")
    }
}

fn delimiter(value: &str, left: bool) -> String {
    let value = match value {
        "{" => "\\{".to_string(), "}" => "\\}".to_string(),
        "|" => if left { "\\lvert" } else { "\\rvert" }.to_string(),
        "‖" => if left { "\\lVert" } else { "\\rVert" }.to_string(),
        "〈" | "⟨" => "\\langle".to_string(), "〉" | "⟩" => "\\rangle".to_string(),
        "" => ".".to_string(), other => other.to_string(),
    };
    format!("{}{} ", if left { "\\left" } else { "\\right" }, value)
}

fn child(node: &XmlNode, index: usize) -> String {
    node.children.get(index).map(convert_node).unwrap_or_default()
}

fn styled_latex(node: &XmlNode, body: String) -> String {
    let variant=node.attribute("mathvariant").unwrap_or_default();
    match variant {
        "bold" => format!("\\mathbf{{{body}}}"),
        "bold-italic" => format!("\\boldsymbol{{{body}}}"),
        "double-struck" => format!("\\mathbb{{{body}}}"),
        "script" => format!("\\mathcal{{{body}}}"),
        "bold-script" => format!("\\boldsymbol{{\\mathcal{{{body}}}}}"),
        "fraktur" => format!("\\mathfrak{{{body}}}"),
        "bold-fraktur" => format!("\\boldsymbol{{\\mathfrak{{{body}}}}}"),
        "sans-serif" => format!("\\mathsf{{{body}}}"),
        "bold-sans-serif" => format!("\\boldsymbol{{\\mathsf{{{body}}}}}"),
        "sans-serif-italic" => format!("\\mathsf{{\\mathit{{{body}}}}}"),
        "sans-serif-bold-italic" => format!("\\boldsymbol{{\\mathsf{{\\mathit{{{body}}}}}}}"),
        "monospace" => format!("\\mathtt{{{body}}}"),
        "italic" if node.name == "mstyle" => format!("\\mathit{{{body}}}"),
        "normal" if node.name == "mstyle" => format!("\\mathrm{{{body}}}"),
        _ => body,
    }
}

fn large_operator(node: &XmlNode) -> bool {
    let symbol=node.children.first().map(XmlNode::all_text).unwrap_or_default();
    matches!(symbol.trim(), "∑"|"∏"|"∐"|"∫"|"∬"|"∭"|"∮"|"⋃"|"⋂"|"⋁"|"⋀")
}

fn convert_node(node: &XmlNode) -> String {
    let body=convert_node_unstyled(node);
    if matches!(node.name.as_str(), "mi"|"mn"|"mo"|"mtext"|"mstyle") { styled_latex(node,body) } else { body }
}
fn convert_node_unstyled(node: &XmlNode) -> String {
    match node.name.as_str() {
        "math" | "mrow" | "mstyle" | "semantics" | "mtd" => children(node),
        "annotation" | "annotation-xml" | "none" | "mprescripts" => String::new(),
        "mspace" => {
            let width=node.attribute("width").unwrap_or("0em");
            if width=="0em" || width=="0" { String::new() } else { format!("\\hspace{{{width}}}") }
        },
        "mi" => {
            let value = node.all_text().trim().to_string();
            let variant = node.attribute("mathvariant").unwrap_or_default().to_ascii_lowercase();
            if (variant.contains("normal") || variant.contains("upright"))
                && value.chars().all(|value| value.is_alphanumeric() || ".,".contains(value))
            {
                format!("\\mathrm{{{}}}", escape_latex(&value, false))
            } else {
                token(&value)
            }
        }
        "mn" | "mo" => token(&node.all_text()),
        "mtext" => {
            let value = node.all_text();
            let trimmed = value.trim();
            if trimmed.chars().count() == 1 && trimmed.chars().all(char::is_alphabetic) {
                format!("\\mathrm{{{}}}", escape_latex(trimmed, false))
            } else {
                format!("\\text{{{}}}", escape_latex(&value, true))
            }
        }
        "mfrac" if matches!(node.attribute("linethickness"), Some("0"|"0pt"|"0px")) => format!("{{{}\\atop {}}}",child(node,0),child(node,1)),
        "mfrac" => format!("\\frac{{{}}}{{{}}}", child(node, 0), child(node, 1)),
        "msqrt" => format!("\\sqrt{{{}}}", children(node)),
        "mroot" => format!("\\sqrt[{}]{{{}}}", child(node, 1), child(node, 0)),
        "msub" => format!("{}_{{{}}}", group(child(node, 0)), child(node, 1)),
        "munder" if large_operator(node) => format!("{}\\limits_{{{}}}",child(node,0),child(node,1)),
        "munder" => convert_accent(node,true),
        "msup" => format!("{}^{{{}}}", group(child(node, 0)), child(node, 1)),
        "munderover" if large_operator(node) => format!("{}\\limits_{{{}}}^{{{}}}",child(node,0),child(node,1),child(node,2)),
        "munderover" => format!("\\overset{{{}}}{{\\underset{{{}}}{{{}}}}}",child(node,2),child(node,1),child(node,0)),
        "msubsup" => format!(
            "{}_{{{}}}^{{{}}}",
            group(child(node, 0)), child(node, 1), child(node, 2)
        ),
        "mover" if large_operator(node) => format!("{}\\limits^{{{}}}",child(node,0),child(node,1)),
        "mover" => convert_accent(node, false),
        "mfenced" => {
            let open = node.attribute("open").unwrap_or("(");
            let close = node.attribute("close").unwrap_or(")");
            let separator = node.attribute("separators").unwrap_or(",");
            format!(
                "{}{}{}",
                delimiter(open, true),
                node.children.iter().map(convert_node).collect::<Vec<_>>().join(separator),
                delimiter(close, false)
            )
        }
        "mtable" => {
            let rows = node
                .children
                .iter()
                .filter(|row| matches!(row.name.as_str(), "mtr" | "mlabeledtr"))
                .map(|row| row.children.iter().map(convert_node).collect::<Vec<_>>().join(" & "))
                .collect::<Vec<_>>();
            format!("\\begin{{matrix}}{}\\end{{matrix}}", rows.join(" \\\\ "))
        }
        "menclose" => {
            let body = children(node);
            let notation = node.attribute("notation").unwrap_or_default().to_ascii_lowercase();
            if notation.contains("box") {
                format!("\\boxed{{{body}}}")
            } else if notation.contains("radical") {
                format!("\\sqrt{{{body}}}")
            } else {
                body
            }
        }
        "mphantom" => format!("\\phantom{{{}}}", children(node)),
        "mmultiscripts" => convert_multiscripts(node),
        _ => {
            let nested = children(node);
            if nested.is_empty() { token(&node.all_text()) } else { nested }
        }
    }
}

fn convert_multiscripts(node: &XmlNode) -> String {
    let Some(base) = node.children.first() else { return String::new(); };
    let split=node.children.iter().position(|n| n.name=="mprescripts").unwrap_or(node.children.len());
    let pairs=|values:&[XmlNode]| values.chunks(2).map(|pair| {
        let sub=convert_node(&pair[0]);
        let sup=pair.get(1).map(convert_node).unwrap_or_default();
        format!("{}{}",if sub.is_empty(){String::new()}else{format!("_{{{sub}}}")},if sup.is_empty(){String::new()}else{format!("^{{{sup}}}")})
    }).collect::<Vec<_>>();
    let mut result=String::new();
    if split<node.children.len() { for pair in pairs(&node.children[split+1..]) { result.push_str(&format!("{{}}{pair}")); } }
    result.push_str(&group(convert_node(base)));
    for (i,pair) in pairs(&node.children[1..split]).iter().enumerate() { if i>0 {result.push_str("{}");} result.push_str(pair); }
    result
}

fn convert_accent(node: &XmlNode, under: bool) -> String {
    let body = child(node, 0);
    let mark = node.children.get(1).map(XmlNode::all_text).unwrap_or_default();
    let accent = node.attribute(if under {"accentunder"} else {"accent"}) != Some("false");
    let command = match mark.trim() {
        "¯" | "̄" | "_" if under && accent => "\\underline",
        "¯" | "̄" if accent => "\\overline", "̂" => "\\hat", "˜" | "̃" => "\\tilde",
        "˙" | "̇" => "\\dot", "¨" | "̈" => "\\ddot", "→" | "⃗" => "\\vec",
        "⌢" => "\\widehat", "⏞" => "\\overbrace", "⏟" => "\\underbrace",
        _ if under => "\\underset", _ => "\\overset",
    };
    if matches!(command, "\\underset" | "\\overset") {
        format!("{command}{{{}}}{{{body}}}", token(mark.trim()))
    } else {
        format!("{command}{{{body}}}")
    }
}

fn collapse_repeated_spaces(value: &str) -> String {
    let mut result = String::with_capacity(value.len());
    let mut previous_space = false;
    for character in value.chars() {
        if character == ' ' {
            if !previous_space {
                result.push(character);
            }
            previous_space = true;
        } else {
            result.push(character);
            previous_space = false;
        }
    }
    result
}

/// Native editing must not silently flatten unhandled MathML structures. The
/// permissive batch converter remains available for the existing import path.
pub fn mathml_to_latex_checked(value: &str) -> Result<String,String> {
    let root=parse_xml(value)?;
    if root.name!="math" { return Err("The native equation did not export a MathML root".into()); }
    fn validate(node:&XmlNode)->Result<(),String> {
        let arity=match node.name.as_str() {
            "math"|"mrow"|"mstyle"|"semantics"|"mtd"|"mtr"|"mlabeledtr"|"mtable"|"msqrt"|"mfenced"|"menclose"|"mphantom"|"mmultiscripts" => None,
            "mi"|"mn"|"mo"|"mtext"|"mspace"|"none"|"mprescripts"=>Some(0),
            "msub"|"msup"|"munder"|"mover"|"mfrac"|"mroot"=>Some(2),
            "msubsup"|"munderover"=>Some(3),
            "annotation"|"annotation-xml"=>return Ok(()),
            other=>return Err(format!("This native equation uses an unsupported MathML element ({other}); the original equation was not changed")),
        };
        if arity.is_some_and(|n|node.children.len()!=n) {return Err(format!("Invalid native MathML structure: {}",node.name));}
        if node.attribute("bevelled")==Some("true") { return Err("Native skewed fractions require a supported lossless conversion before editing; the original was not changed".into()); }
        if let Some(v)=node.attribute("mathvariant") {
            if !matches!(v,"normal"|"upright"|"italic"|"bold"|"bold-italic"|"double-struck"|"script"|"bold-script"|"fraktur"|"bold-fraktur"|"sans-serif"|"bold-sans-serif"|"sans-serif-italic"|"sans-serif-bold-italic"|"monospace") {
                return Err(format!("Unsupported native mathematical style: {v}"));
            }
        }
        for child in &node.children {validate(child)?;} Ok(())
    }
    validate(&root)?;
    let result=convert_node(&root).trim().to_string();
    if result.is_empty() {Err("The native equation is empty".into())} else {Ok(result)}
}

pub fn mathml_to_latex(value: &str) -> Result<String, String> {
    let root = parse_xml(value)?;
    let result = collapse_repeated_spaces(&convert_node(&root))
        .trim()
        .to_string();
    if result.is_empty() {
        Err("Word MathML did not contain a recoverable formula".to_string())
    } else {
        Ok(result)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_round_trip_preserves_styles_entities_and_grouping() {
        let styled=r#"<math><mstyle mathvariant="bold"><mi>x</mi><mi>y</mi></mstyle><mi mathvariant="double-struck">R</mi><mtext>A &amp; B</mtext></math>"#;
        assert_eq!(mathml_to_latex_checked(styled).unwrap(), r"\mathbf{xy}\mathbb{R}\text{A & B}".replace('&',r"\&"));
        assert_eq!(mathml_to_latex_checked(r"<math><msup><mrow><mi>α</mi><mo>+</mo><mi>x</mi></mrow><mn>2</mn></msup></math>").unwrap(), r"{\alpha +x}^{2}");
        assert!(mathml_to_latex_checked("<math><unknown><mi>x</mi></unknown></math>").is_err());
    }
    #[test]
    fn native_round_trip_preserves_under_accents_and_prescripts() {
        assert_eq!(mathml_to_latex_checked("<math><munder accentunder=\"true\"><mi>x</mi><mo>¯</mo></munder></math>").unwrap(),r"\underline{x}");
        assert_eq!(mathml_to_latex_checked("<math><mmultiscripts><mi>X</mi><mi>i</mi><none/><mprescripts/><mi>j</mi><mn>2</mn></mmultiscripts></math>").unwrap(),r"{}_{j}^{2}X_{i}");
        assert_eq!(mathml_to_latex_checked("<math><mover accent=\"false\"><mi>x</mi><mi>n</mi></mover></math>").unwrap(),r"\overset{n}{x}");
    }
    #[test]
    fn converts_structured_mathml() {
        let value = r#"<math xmlns="http://www.w3.org/1998/Math/MathML"><mfrac><mi>a</mi><msup><mi>b</mi><mn>2</mn></msup></mfrac></math>"#;
        assert_eq!(mathml_to_latex(value).unwrap(), "\\frac{a}{b^{2}}");
    }
}
