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
    if values.get("sourceKind").map(String::as_str) == Some("nativeNumber") {
        return build_native_numbers(&values, output);
    }
    if values.get("sourceKind").map(String::as_str) == Some("imageLetterFonts") {
        return build_image_letter_fonts(&values, output);
    }
    if values.get("sourceKind").map(String::as_str) == Some("displayNumbering") {
        return build_display_numbering(&values, output);
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
// Equation format v2 (docs/equation-format-v2.md): the Word add-in decides a
// number's context (label, heading prefix, restart) from the document and asks
// the app for the finished numbered equations, built by the shared writer
// (WordOmmlConverter.BuildWordNativeNumberedOmml on Windows). One package holds
// every requested equation, one paragraph each; the add-in inserts it or
// copies the number slots out of it.
//
// Request rows: sessionId, sourceDocumentId, itemCount, mathFont (Base64URL),
// optional ommlBase64 (the equation body; Windows' x placeholder otherwise),
// and per item N: itemNSequenceName (Base64URL), itemNHeadingLevel (0-2),
// itemNSeparator (Base64URL), optional itemNLiteralPrefix (Base64URL; absent =
// STYLEREF when the level is not 0), optional itemNRestart, itemNInitial.
const NATIVE_NUMBER_DOCX_FILE: &str = "native-number.docx";
const NATIVE_NUMBER_PLACEHOLDER_BODY: &str = "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>";

fn build_native_numbers(values: &HashMap<String, String>, output: &Path) -> Result<(), String> {
    let session_id = required(values, "sessionId")?;
    let source_document_id = required(values, "sourceDocumentId")?;
    let item_count = required(values, "itemCount")?
        .parse::<usize>()
        .map_err(|_| "The equation number item count is invalid".to_string())?;
    if item_count == 0 || item_count > MAX_ITEMS {
        return Err(format!("The equation number request must contain 1 to {MAX_ITEMS} items"));
    }
    let math_font = decode_utf8(required(values, "mathFont")?, "the Word math font")?;
    let body = match values.get("ommlBase64") {
        Some(value) => decode_utf8(value, "the equation OMML")?,
        None => NATIVE_NUMBER_PLACEHOLDER_BODY.to_string(),
    };
    let mut equations = Vec::with_capacity(item_count);
    for index in 1..=item_count {
        let context = native_number_context(values, index)?;
        equations.push(native_number_equation(&body, &context, &math_font)?);
    }
    let package = super::omml_write::word_equation_docx(&equations, &math_font)?;
    let docx_path = output.with_file_name(NATIVE_NUMBER_DOCX_FILE);
    fs::write(&docx_path, &package)
        .map_err(|error| format!("Unable to write {}: {error}", docx_path.display()))?;
    let rows = [
        format!("protocolVersion={PROTOCOL_VERSION}"),
        format!("sessionId={session_id}"),
        "operation=nativeNumber".to_string(),
        "outputKind=docx".to_string(),
        format!("sourceDocumentId={source_document_id}"),
        "bookmarkName=-".to_string(),
        format!("itemCount={item_count}"),
        format!("docxPath={}", docx_path.display()),
    ];
    fs::write(output, rows.join("\n"))
        .map_err(|error| format!("Unable to write {}: {error}", output.display()))
}

/// The number of request item N: SEQ label, heading level and separator, the
/// literal heading prefix (None = STYLEREF), the \r value and the initial
/// result.
struct NativeNumberContext {
    sequence_name: String,
    heading_level: u32,
    separator: String,
    literal_prefix: Option<String>,
    restart: Option<u32>,
    initial: String,
}

fn native_number_context(
    values: &HashMap<String, String>,
    index: usize,
) -> Result<NativeNumberContext, String> {
    let key = |name: &str| format!("item{index}{name}");
    let sequence_name = decode_utf8(required(values, &key("SequenceName"))?, "the SEQ label")?;
    let heading_level = required(values, &key("HeadingLevel"))?
        .parse::<u32>()
        .ok()
        .filter(|level| *level <= 2)
        .ok_or_else(|| "The equation number heading level is invalid".to_string())?;
    let separator = decode_utf8(required(values, &key("Separator"))?, "the number separator")?;
    let literal_prefix = values
        .get(&key("LiteralPrefix"))
        .map(|value| decode_utf8(value, "the number prefix"))
        .transpose()?;
    let restart = values
        .get(&key("Restart"))
        .map(|value| {
            value
                .parse::<u32>()
                .ok()
                .filter(|restart| *restart >= 1)
                .ok_or_else(|| "The equation number restart is invalid".to_string())
        })
        .transpose()?;
    let initial = required(values, &key("Initial"))?;
    if initial.is_empty() || initial.len() > 9 || !initial.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err("The equation number initial result is invalid".to_string());
    }
    Ok(NativeNumberContext {
        sequence_name,
        heading_level,
        separator,
        literal_prefix,
        restart,
        initial: initial.to_string(),
    })
}

/// One numbered equation, as Word imports it: the body numbered by the shared
/// writer, in the document's math font.
fn native_number_equation(
    body: &str,
    context: &NativeNumberContext,
    math_font: &str,
) -> Result<String, String> {
    let numbered = super::omml_write::build_word_native_numbered_omml(
        body,
        &context.sequence_name,
        None,
        context.heading_level,
        &context.separator,
        &context.initial,
        context.literal_prefix.as_deref(),
        context.restart,
    )?;
    super::omml_write::insertion_omml(&numbered, math_font)
}

// Batch numbering of display formulas (Windows' Number Display Formulas): one
// request numbers every display formula the add-in found. Each OMML equation
// is numbered with its own body, exactly as Windows' batch reads it
// (ExtractSingleOMath of Word's export) and numbers it
// (BuildWordNativeNumberedOmml); the package holds one paragraph per item in
// item order, and an item that cannot be numbered leaves its paragraph empty
// and reports itemNError. Each VisualTeX image formula's metadata is marked
// numbered, so a later conversion sees what its title says; the metadata is
// written to its own file because it can exceed what Word reads back through
// AppleScriptTask.
//
// Request rows: sessionId, sourceDocumentId, itemCount and imageCount (either
// may be 0). With OMML items: mathFont and ommlRoot (Base64URL; the
// w:document start tag of Word's export, with its namespace declarations), and
// per item N the nativeNumber rows plus itemNOmml, the item's m:oMath as Word
// exported it with "%", CR and LF percent-escaped (and itemNOmmlRoot when that
// export's root differs from ommlRoot). Per image N: imageNMetadata.
const DISPLAY_NUMBERING_DOCX_FILE: &str = "display-numbering.docx";
const DISPLAY_NUMBERING_IMAGES_FILE: &str = "display-numbering-images.txt";

fn build_display_numbering(values: &HashMap<String, String>, output: &Path) -> Result<(), String> {
    let session_id = required(values, "sessionId")?;
    let source_document_id = required(values, "sourceDocumentId")?;
    let count = |key: &str| {
        required(values, key)?
            .parse::<usize>()
            .ok()
            .filter(|count| *count <= MAX_ITEMS)
            .ok_or_else(|| format!("The display numbering {key} is invalid"))
    };
    let item_count = count("itemCount")?;
    let image_count = count("imageCount")?;
    if item_count + image_count == 0 {
        return Err("The display numbering request contains no formula".to_string());
    }
    let mut rows = vec![
        format!("protocolVersion={PROTOCOL_VERSION}"),
        format!("sessionId={session_id}"),
        "operation=displayNumbering".to_string(),
        "outputKind=displayNumbering".to_string(),
        format!("sourceDocumentId={source_document_id}"),
        "bookmarkName=-".to_string(),
        format!("itemCount={item_count}"),
        format!("imageCount={image_count}"),
    ];

    if item_count > 0 {
        let math_font = decode_utf8(required(values, "mathFont")?, "the Word math font")?;
        let root = document_root(required(values, "ommlRoot")?)?;
        let mut equations = Vec::with_capacity(item_count);
        for index in 1..=item_count {
            let context = native_number_context(values, index)?;
            let exported = percent_decode(required(values, &format!("item{index}Omml"))?)?;
            // An export whose root differs from the first one names its own.
            let own_root = values
                .get(&format!("item{index}OmmlRoot"))
                .map(|value| document_root(value))
                .transpose()?;
            let numbered = super::omml_write::extract_single_omath(&format!(
                "{}<w:body><w:p>{exported}</w:p></w:body></w:document>",
                own_root.as_deref().unwrap_or(&root)
            ))
            .and_then(|body| native_number_equation(&body, &context, &math_font));
            match numbered {
                Ok(equation) => equations.push(equation),
                Err(error) => {
                    equations.push(String::new());
                    rows.push(format!(
                        "item{index}Error={}",
                        URL_SAFE_NO_PAD.encode(error.as_bytes())
                    ));
                }
            }
        }
        let package = super::omml_write::word_equation_docx(&equations, &math_font)?;
        let docx_path = output.with_file_name(DISPLAY_NUMBERING_DOCX_FILE);
        fs::write(&docx_path, &package)
            .map_err(|error| format!("Unable to write {}: {error}", docx_path.display()))?;
        rows.push(format!("docxPath={}", docx_path.display()));
    }

    if image_count > 0 {
        let mut image_rows = vec![
            format!("protocolVersion={PROTOCOL_VERSION}"),
            format!("sessionId={session_id}"),
            "operation=displayNumberingImages".to_string(),
            "outputKind=metadata".to_string(),
            format!("sourceDocumentId={source_document_id}"),
            "bookmarkName=-".to_string(),
            format!("itemCount={image_count}"),
            format!("imageCount={image_count}"),
        ];
        for index in 1..=image_count {
            let metadata = required(values, &format!("image{index}Metadata"))?;
            match super::macos_offline::numbered_formula_metadata(metadata) {
                Ok(numbered) => {
                    image_rows.push(format!("image{index}Metadata={numbered}"));
                    // The new number takes the formula's letter font.
                    if let Some(font) = formula_letter_font(metadata) {
                        image_rows.push(format!("image{index}LetterFont={font}"));
                    }
                }
                Err(error) => image_rows.push(format!(
                    "image{index}Error={}",
                    URL_SAFE_NO_PAD.encode(error.as_bytes())
                )),
            }
        }
        let images_path = output.with_file_name(DISPLAY_NUMBERING_IMAGES_FILE);
        write_output_manifest(&images_path, &image_rows)?;
        rows.push(format!("imagesPath={}", images_path.display()));
    }
    write_output_manifest(output, &rows)
}

/// The letter font an image formula's metadata records, if it is one Word
/// numbers can follow.
fn formula_letter_font(metadata: &str) -> Option<String> {
    super::macos_offline::decode_metadata(metadata)
        .ok()?
        .formula_letter_font
        .filter(|font| {
            matches!(
                font.as_str(),
                "katex" | "times" | "cambria" | "stix" | "palatino" | "helvetica"
            )
        })
}

/// The letter fonts of image formulas whose numbers are written or resized
/// (formulas from before the add-in recorded them). An unreadable metadata
/// string or one without a letter font gets no row.
fn build_image_letter_fonts(values: &HashMap<String, String>, output: &Path) -> Result<(), String> {
    let session_id = required(values, "sessionId")?;
    let source_document_id = required(values, "sourceDocumentId")?;
    let image_count = required(values, "imageCount")?
        .parse::<usize>()
        .ok()
        .filter(|count| (1..=MAX_ITEMS).contains(count))
        .ok_or_else(|| "The image letter font imageCount is invalid".to_string())?;
    let mut rows = vec![
        format!("protocolVersion={PROTOCOL_VERSION}"),
        format!("sessionId={session_id}"),
        "operation=imageLetterFonts".to_string(),
        "outputKind=imageLetterFonts".to_string(),
        // The rows every reply the add-in parses carries.
        format!("sourceDocumentId={source_document_id}"),
        "bookmarkName=-".to_string(),
        format!("itemCount={image_count}"),
        format!("imageCount={image_count}"),
    ];
    for index in 1..=image_count {
        if let Some(font) = formula_letter_font(required(values, &format!("image{index}Metadata"))?) {
            rows.push(format!("image{index}LetterFont={font}"));
        }
    }
    write_output_manifest(output, &rows)
}

/// The w:document start tag of Word's export (Base64URL), with the namespace
/// declarations its equations use.
fn document_root(value: &str) -> Result<String, String> {
    let root = decode_utf8(value, "the Word document element")?;
    if !root.starts_with("<w:document")
        || !root.ends_with('>')
        || root.ends_with("/>")
        || root[1..].contains('<')
    {
        return Err("The Word document element of the equations is invalid".to_string());
    }
    Ok(root)
}

/// Undoes the add-in's escaping of one manifest value: "%XX" is the byte XX.
fn percent_decode(value: &str) -> Result<String, String> {
    let bytes = value.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'%' {
            let byte = value
                .get(index + 1..index + 3)
                .and_then(|hex| u8::from_str_radix(hex, 16).ok())
                .ok_or_else(|| "An exported equation contains an invalid escape".to_string())?;
            decoded.push(byte);
            index += 3;
        } else {
            decoded.push(bytes[index]);
            index += 1;
        }
    }
    String::from_utf8(decoded).map_err(|_| "An exported equation is not valid UTF-8".to_string())
}

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

    #[test]
    fn percent_decode_restores_escaped_values() {
        assert_eq!(percent_decode("a%25b%0D%0Ac").unwrap(), "a%b\r\nc");
        assert_eq!(percent_decode("x\u{3b1}").unwrap(), "x\u{3b1}");
        assert!(percent_decode("bad%2").is_err());
        assert!(percent_decode("bad%zz").is_err());
    }

    // A Word export (the w:document start tag and one m:oMath) and the
    // metadata of a VisualTeX image display formula, from a Mac batch import.
    const WORD_ROOT: &str = r#"<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">"#;
    const WORD_EQUATION: &str = r#"<m:oMath><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>x</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>=</m:t></m:r><m:f><m:fPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></m:ctrlPr></m:fPr><m:num><m:r><m:rPr><m:sty m:val="p"/></m:rPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>7</m:t></m:r></m:num><m:den><m:r><m:rPr><m:sty m:val="p"/></m:rPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>6</m:t></m:r></m:den></m:f></m:oMath>"#;
    const IMAGE_METADATA: &str = "visualtex:v1:deflate:pZJNb9QwEIb_SuVzHNlO4k0icYBKFStRqaKICgiLvPGksTZxVo4XtkT574xLdrcXuHQOPvidj2deeyJj3UKvSEl-mvGgOg9H2gyuP3SKRIv4GdxoBktKHpFFW2us0HwlskantFCpomnCt7TQoqZC11qmiZQrnWMTb3wHmL7u94PzoK-0Gfederq6zOkUzsWUqjLW_5jYvJnEfNxMxVxVUVX1yreun_R8fFNVjVP1JDYTZ_McjlBuLIyk_DYRE7iKOuMamKQ5C1wyFXSbCEFXXHHWrLQWUr926Pw9IvWg4QZ3UB67OPULey6r3aKCd9tuqHd4a2wgXPfqEW6x671_ejYEx3tU7aHfggMkb1Q3QkQcWA3uwWjf3iEgT9I4wSXkKbJTynswj60PORmLuUhSeYnwVNbfm99wh3hcxOz8eB_Ae3A3KCPE7tmFs3bdBi_hLBpvSBjXIKGt4S9UaMhZnGcsEy_EBQfVlMdJxnP5QnynRgguBJnyLJZFIiQ7RYomBXvWdncNFum-fFTeDKRkMW5bO0BK_WB8e_6LhMciDt_rsNf_FpfKt2EZwYSknFFefGJZyfIyKeKkEF8vPf6fNv8B";

    fn escape_for_manifest(value: &str) -> String {
        value.replace('%', "%25").replace('\r', "%0D").replace('\n', "%0A")
    }

    fn run_manifest(rows: &[String]) -> (PathBuf, HashMap<String, String>) {
        let directory = std::env::temp_dir().join(format!(
            "visualtex-display-numbering-{}",
            uuid::Uuid::new_v4()
        ));
        fs::create_dir_all(&directory).unwrap();
        let input = directory.join("formula-restore-source.txt");
        let output = directory.join("omml-latex-result.txt");
        fs::write(&input, rows.join("\n")).unwrap();
        convert_manifest(&input, &output).unwrap();
        let result = fs::read_to_string(&output)
            .unwrap()
            .lines()
            .filter_map(|row| row.split_once('='))
            .map(|(key, value)| (key.to_string(), value.to_string()))
            .collect();
        (directory, result)
    }

    fn base64(value: &str) -> String {
        URL_SAFE_NO_PAD.encode(value.as_bytes())
    }

    fn display_item(index: usize, omml: &str, initial: u32) -> Vec<String> {
        vec![
            format!("item{index}SequenceName={}", base64("Equation")),
            format!("item{index}HeadingLevel=1"),
            format!("item{index}Separator={}", base64(".")),
            format!("item{index}LiteralPrefix={}", base64("4")),
            format!("item{index}Initial={initial}"),
            format!("item{index}Omml={}", escape_for_manifest(omml)),
        ]
    }

    #[test]
    fn display_numbering_numbers_each_equation_with_its_own_body() {
        use std::io::Read as _;
        let second = WORD_EQUATION.replace("<m:t>x</m:t>", "<m:t>y</m:t>");
        let mut rows = vec![
            "protocolVersion=1".to_string(),
            "sessionId=12345678-1234-4234-9234-123456789abc".to_string(),
            "sourceKind=displayNumbering".to_string(),
            format!("sourceDocumentId={}", base64("Macintosh HD:test.docx")),
            "bookmarkName=-".to_string(),
            "itemCount=3".to_string(),
            "imageCount=0".to_string(),
            format!("mathFont={}", base64("Cambria Math")),
            format!("ommlRoot={}", base64(WORD_ROOT)),
        ];
        rows.extend(display_item(1, WORD_EQUATION, 1));
        rows.extend(display_item(2, &second, 2));
        // Two equations in one export cannot be numbered as one.
        rows.extend(display_item(3, &format!("{WORD_EQUATION}{WORD_EQUATION}"), 3));
        let (directory, result) = run_manifest(&rows);
        assert_eq!(result["operation"], "displayNumbering");
        assert_eq!(result["itemCount"], "3");
        assert!(!result.contains_key("item1Error") && !result.contains_key("item2Error"));
        let error = URL_SAFE_NO_PAD.decode(&result["item3Error"]).unwrap();
        assert!(!String::from_utf8(error).unwrap().is_empty());
        assert!(!result.contains_key("imagesPath"));

        let package = fs::read(&result["docxPath"]).unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(package)).unwrap();
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut document)
            .unwrap();
        let paragraphs: Vec<&str> = document.split("<w:p>").skip(1).collect();
        assert_eq!(paragraphs.len(), 3);
        for (paragraph, letter, initial) in [(paragraphs[0], "x", "1"), (paragraphs[1], "y", "2")] {
            // Serialized as Windows' writer serializes it: the math namespace
            // as the default one.
            assert!(paragraph.starts_with("<oMath xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"), "{paragraph}");
            assert!(paragraph.contains("<eqArr>"), "{paragraph}");
            assert!(paragraph.contains(&format!("<t>{letter}</t>")));
            assert!(paragraph.contains("<t>#</t>"));
            assert!(paragraph.contains(" SEQ Equation \\s 1 \\* ARABIC \\* MERGEFORMAT "));
            assert!(paragraph.contains(&format!("<t>{initial}</t>")));
            assert!(paragraph.contains("<t>4</t>") && paragraph.contains("<t>.</t>"));
            assert!(paragraph.contains("<f>") && paragraph.contains("Cambria Math"));
        }
        assert!(!paragraphs[2].contains("oMath"));
        fs::remove_dir_all(directory).unwrap();
    }

    #[test]
    fn display_numbering_marks_image_metadata_numbered() {
        let rows = vec![
            "protocolVersion=1".to_string(),
            "sessionId=12345678-1234-4234-9234-123456789abc".to_string(),
            "sourceKind=displayNumbering".to_string(),
            format!("sourceDocumentId={}", base64("Macintosh HD:test.docx")),
            "bookmarkName=-".to_string(),
            "itemCount=0".to_string(),
            "imageCount=2".to_string(),
            format!("image1Metadata={IMAGE_METADATA}"),
            "image2Metadata=visualtex:v1:deflate:broken".to_string(),
        ];
        let (directory, result) = run_manifest(&rows);
        assert_eq!(result["imageCount"], "2");
        assert!(!result.contains_key("docxPath"));
        let images: HashMap<String, String> = fs::read_to_string(&result["imagesPath"])
            .unwrap()
            .lines()
            .filter_map(|row| row.split_once('='))
            .map(|(key, value)| (key.to_string(), value.to_string()))
            .collect();
        assert_eq!(images["imageCount"], "2");
        assert_eq!(images["operation"], "displayNumberingImages");
        assert_eq!(images["itemCount"], "2");
        let before = super::super::macos_offline::decode_metadata(IMAGE_METADATA).unwrap();
        let after = super::super::macos_offline::decode_metadata(&images["image1Metadata"]).unwrap();
        assert!(!before.numbered && after.numbered);
        assert_eq!(after.formula_id, before.formula_id);
        assert_eq!(after.latex, before.latex);
        assert_eq!(after.display_mode, "block");
        assert_eq!(after.reference_width_pt, before.reference_width_pt);
        assert_eq!(after.updated_at, before.updated_at);
        assert!(images.contains_key("image2Error") && !images.contains_key("image2Metadata"));
        fs::remove_dir_all(directory).unwrap();
    }

    #[test]
    fn image_letter_fonts_come_from_each_formula_metadata() {
        let rows = vec![
            "protocolVersion=1".to_string(),
            "sessionId=12345678-1234-4234-9234-123456789abc".to_string(),
            "sourceKind=imageLetterFonts".to_string(),
            format!("sourceDocumentId={}", base64("Macintosh HD:test.docx")),
            "bookmarkName=-".to_string(),
            "imageCount=2".to_string(),
            format!("image1Metadata={IMAGE_METADATA}"),
            "image2Metadata=visualtex:v1:deflate:broken".to_string(),
        ];
        let (directory, result) = run_manifest(&rows);
        assert_eq!(result["operation"], "imageLetterFonts");
        assert_eq!(result["imageCount"], "2");
        // VTParseDocumentImportManifestText requires these.
        for key in ["protocolVersion", "sessionId", "outputKind", "sourceDocumentId", "bookmarkName", "itemCount"] {
            assert!(result.contains_key(key), "reply lacks {key}");
        }
        let expected = super::super::macos_offline::decode_metadata(IMAGE_METADATA)
            .unwrap()
            .formula_letter_font;
        assert_eq!(result.get("image1LetterFont").cloned(), expected);
        assert!(!result.contains_key("image2LetterFont"));
        fs::remove_dir_all(directory).unwrap();
    }

    #[test]
    fn formula_letter_font_accepts_only_known_fonts() {
        let mut metadata = super::super::macos_offline::decode_metadata(IMAGE_METADATA).unwrap();
        for (font, expected) in [("times", Some("times")), ("comic", None)] {
            metadata.formula_letter_font = Some(font.to_string());
            let encoded = super::super::macos_offline::encode_metadata(&metadata).unwrap();
            assert_eq!(formula_letter_font(&encoded).as_deref(), expected);
        }
        assert_eq!(formula_letter_font("visualtex:v1:deflate:broken"), None);
    }

    #[test]
    fn display_numbering_rejects_an_empty_or_oversized_request() {
        let rows = |items: &str, images: &str| {
            vec![
                "protocolVersion=1".to_string(),
                "sessionId=12345678-1234-4234-9234-123456789abc".to_string(),
                "sourceKind=displayNumbering".to_string(),
                format!("sourceDocumentId={}", base64("doc")),
                "bookmarkName=-".to_string(),
                format!("itemCount={items}"),
                format!("imageCount={images}"),
            ]
        };
        for (items, images) in [("0", "0"), ("1001", "0"), ("0", "x")] {
            let values: HashMap<String, String> = rows(items, images)
                .iter()
                .filter_map(|row| row.split_once('='))
                .map(|(key, value)| (key.to_string(), value.to_string()))
                .collect();
            assert!(build_display_numbering(&values, Path::new("/nonexistent/out.txt")).is_err());
        }
    }
}
