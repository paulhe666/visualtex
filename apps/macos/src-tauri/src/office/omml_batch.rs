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
        let math_ml = word_omml_to_mathml(&omml, display_mode == "block")?;
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

/// Native editing must not silently flatten unhandled MathML structures: the
/// equation is checked before it is converted with the shared Word reader.
pub fn mathml_to_latex_checked(value: &str) -> Result<String,String> {
    let root=parse_xml(value)?;
    if root.name!="math" { return Err("The native equation did not export a MathML root".into()); }
    fn validate(node:&XmlNode)->Result<(),String> {
        let arity=match node.name.as_str() {
            "math"|"mrow"|"mstyle"|"mpadded"|"semantics"|"mtd"|"mtr"|"mlabeledtr"|"mtable"|"msqrt"|"mfenced"|"menclose"|"mphantom"|"mmultiscripts" => None,
            // Equation-array alignment points read back as the `&` of `aligned`.
            "mi"|"mn"|"mo"|"mtext"|"mspace"|"none"|"mprescripts"|"maligngroup"|"malignmark"=>Some(0),
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
    let result=super::omml_latex::mathml_to_word_equation_latex(value)?.trim().to_string();
    if result.is_empty() {Err("The native equation is empty".into())} else {Ok(result)}
}

/// Word MathML as LaTeX, read the same way as on Windows.
pub fn mathml_to_latex(value: &str) -> Result<String, String> {
    let result = super::omml_latex::mathml_to_word_equation_latex(value)?.trim().to_string();
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
    fn native_edit_checks_the_structure_before_converting() {
        assert!(mathml_to_latex_checked("<math><unknown><mi>x</mi></unknown></math>").is_err());
        assert!(mathml_to_latex_checked("<math><mfrac bevelled=\"true\"><mi>a</mi><mi>b</mi></mfrac></math>").is_err());
        assert!(mathml_to_latex_checked("<math><msub><mi>a</mi></msub></math>").is_err());
        assert_eq!(mathml_to_latex_checked("<math><munder accentunder=\"true\"><mi>x</mi><mo>¯</mo></munder></math>").unwrap(),r"\underline{x}");
        assert_eq!(
            mathml_to_latex_checked("<math><mtable><mtr><mtd><mrow><maligngroup/><mi>a</mi><malignmark/><mo>=</mo><mi>b</mi></mrow></mtd></mtr></mtable></math>").unwrap(),
            r"\begin{aligned}a&=b\end{aligned}"
        );
    }
    #[test]
    fn converts_structured_mathml() {
        let value = r#"<math xmlns="http://www.w3.org/1998/Math/MathML"><mfrac><mi>a</mi><msup><mi>b</mi><mn>2</mn></msup></mfrac></math>"#;
        assert_eq!(mathml_to_latex(value).unwrap(), "\\frac{a}{b^2}");
    }
}
