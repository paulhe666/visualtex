//! Word `VisualTeX.Formula.1` OLE objects, interchangeable with the Windows
//! LocalServer.
//!
//! Word for Mac cannot host an OLE server, but a document can still carry the
//! same embedded object: VisualTeX writes the storage that the Windows server
//! persists, and reads it back when the user edits the formula. The contract is
//! defined by `apps/windows/src-windows/VisualTeX.FormulaOleServer`
//! (`FormulaOleContract.h`, `FormulaOleObject.rgs`, `FormulaOleObject.cpp`).

use crate::office::sessions::VisualTeXFormulaMetadata;
use crate::office::stored_zip::build_stored_zip;
use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};
use serde_json::{Map, Value};
use std::collections::HashMap;
use std::io::{Cursor, Read, Write};
use uuid::Uuid;

pub const PROG_ID: &str = "VisualTeX.Formula.1";
/// `{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}`, the published class identity.
pub const CLSID: Uuid = Uuid::from_u128(0x8FF7F5AA_0D60_48D5_ADBD_65A64B4C827B);
/// Default value of the CLSID key in `FormulaOleObject.rgs`.
const USER_TYPE: &str = "VisualTeX Formula";
const METADATA_STREAM: &str = "VisualTeX.Formula.json";
const EMF_STREAM: &str = "VisualTeX.Preview.emf";
const PNG_STREAM: &str = "VisualTeX.Preview.png";
const COMP_OBJ_STREAM: &str = "\u{1}CompObj";
const OLE_STREAM: &str = "\u{1}Ole";
const OBJ_INFO_STREAM: &str = "\u{3}ObjInfo";
const CF_ENHMETAFILE: u16 = 14;
const PNG_SIGNATURE: &[u8] = b"\x89PNG\r\n\x1a\n";

/// The payload of one VisualTeX OLE object.
#[derive(Debug, Clone, PartialEq)]
pub struct OleFormula {
    /// `VisualTeX.Formula.json`. Kept as a JSON object so that fields written
    /// by newer Windows builds survive a macOS edit.
    pub metadata: Map<String, Value>,
    pub emf: Vec<u8>,
    pub png: Vec<u8>,
}

/// Builds the embedded storage (`word/embeddings/oleObjectN.bin`).
///
/// Besides the server's three streams it contains the `\1CompObj`, `\1Ole`
/// and `\3ObjInfo` streams that Word writes for an embedded object, so the
/// storage is complete before Word ever saves it.
pub fn build_ole_storage(formula: &OleFormula) -> Result<Vec<u8>, String> {
    let metadata = serde_json::to_vec(&formula.metadata)
        .map_err(|error| format!("Unable to serialize OLE formula metadata: {error}"))?;
    let mut file = cfb::CompoundFile::create_with_version(cfb::Version::V3, Cursor::new(Vec::new()))
        .map_err(storage_error)?;
    file.set_storage_clsid("/", CLSID).map_err(storage_error)?;
    for (name, contents) in [
        (COMP_OBJ_STREAM, comp_obj_stream()),
        (OLE_STREAM, ole_stream()),
        (OBJ_INFO_STREAM, obj_info_stream()),
        (METADATA_STREAM, metadata),
        (EMF_STREAM, formula.emf.clone()),
        (PNG_STREAM, formula.png.clone()),
    ] {
        let mut stream = file.create_stream(format!("/{name}")).map_err(storage_error)?;
        stream.write_all(&contents).map_err(storage_error)?;
    }
    file.flush().map_err(storage_error)?;
    Ok(file.into_inner().into_inner())
}

/// Reads an embedded storage written by VisualTeX on either platform.
///
/// This is the trust boundary for formulas read out of a Word document, so it
/// applies the same acceptance rules as the Windows server's
/// `IPersistStorage::Load`.
pub fn read_ole_storage(bytes: &[u8]) -> Result<OleFormula, String> {
    let mut file = cfb::CompoundFile::open(Cursor::new(bytes))
        .map_err(|error| format!("The formula object is not an OLE storage: {error}"))?;
    if *file.root_entry().clsid() != CLSID {
        return Err("The OLE object is not a VisualTeX formula".to_string());
    }
    let mut read_stream = |name: &str| -> Result<Vec<u8>, String> {
        let mut stream = file
            .open_stream(format!("/{name}"))
            .map_err(|_| format!("The VisualTeX formula object has no {name} stream"))?;
        let mut contents = Vec::new();
        stream.read_to_end(&mut contents).map_err(storage_error)?;
        Ok(contents)
    };
    let metadata = read_stream(METADATA_STREAM)?;
    let emf = read_stream(EMF_STREAM)?;
    let png = read_stream(PNG_STREAM)?;
    let metadata = match serde_json::from_slice(&metadata) {
        Ok(Value::Object(map)) => map,
        _ => return Err("The VisualTeX formula metadata is not a JSON object".to_string()),
    };
    if !metadata.contains_key("schemaVersion") || !metadata.contains_key("formulaId") {
        return Err("The VisualTeX formula metadata has no schemaVersion or formulaId".to_string());
    }
    if emf.get(40..44) != Some(&0x464D_4520_u32.to_le_bytes()[..]) || !png.starts_with(PNG_SIGNATURE) {
        return Err("The VisualTeX formula object has an invalid preview".to_string());
    }
    Ok(OleFormula { metadata, emf, png })
}

/// The metadata stored in the OLE object after a macOS create or edit.
///
/// Starts from the object's current JSON (if any) so that fields this build
/// does not know about are carried forward, then overwrites every field the
/// session owns. `baseline_px` is the export baseline from the top edge, in
/// the same CSS pixels as `renderHeightPx`.
pub fn ole_metadata(
    original: Option<&Map<String, Value>>,
    current: &VisualTeXFormulaMetadata,
    baseline_px: Option<f64>,
) -> Result<Map<String, Value>, String> {
    let mut metadata = original.cloned().unwrap_or_default();
    match serde_json::to_value(current) {
        Ok(Value::Object(fields)) => metadata.extend(fields),
        _ => return Err("Unable to serialize the formula metadata".to_string()),
    }
    if let Some(baseline) = baseline_px {
        metadata.insert("baseline".to_string(), Value::from(baseline));
    }
    Ok(metadata)
}

/// Converts OLE metadata into the session metadata used by the editor.
pub fn session_metadata(metadata: &Map<String, Value>) -> Result<VisualTeXFormulaMetadata, String> {
    serde_json::from_value(Value::Object(metadata.clone()))
        .map_err(|error| format!("The VisualTeX formula metadata is incomplete: {error}"))
}

/// A minimal DOCX whose body is one inline `VisualTeX.Formula.1` object.
///
/// Word for Mac imports it with `Range.InsertFile`/`FormattedText`, the same
/// staging route as the SVG picture formulas. `object_id` becomes the VML
/// `ObjectID`; Word renumbers it if it collides in the target document.
pub fn build_word_ole_docx(
    storage: &[u8],
    emf: &[u8],
    width_pt: f64,
    height_pt: f64,
    object_id: u32,
) -> Result<Vec<u8>, String> {
    if !(width_pt.is_finite() && height_pt.is_finite() && width_pt > 0.0 && height_pt > 0.0) {
        return Err("Word OLE staging dimensions must be positive".to_string());
    }
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Default Extension="emf" ContentType="image/x-emf"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let package_relationships = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
    let document_relationships = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.emf"/>
  <Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>
</Relationships>"#;
    let width_twips = (width_pt * 20.0).round() as i64;
    let height_twips = (height_pt * 20.0).round() as i64;
    let width = format_points(width_pt);
    let height = format_points(height_pt);
    // The shapetype is Word's standard picture frame (`_x0000_t75`) that it
    // writes around every OLE presentation.
    let document = format!(
        r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
  <w:body>
    <w:p>
      <w:r>
        <w:object w:dxaOrig="{width_twips}" w:dyaOrig="{height_twips}">
          <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
            <v:stroke joinstyle="miter"/>
            <v:formulas>
              <v:f eqn="if lineDrawn pixelLineWidth 0"/>
              <v:f eqn="sum @0 1 0"/>
              <v:f eqn="sum 0 0 @1"/>
              <v:f eqn="prod @2 1 2"/>
              <v:f eqn="prod @3 21600 pixelWidth"/>
              <v:f eqn="prod @3 21600 pixelHeight"/>
              <v:f eqn="sum @0 0 1"/>
              <v:f eqn="prod @6 1 2"/>
              <v:f eqn="prod @7 21600 pixelWidth"/>
              <v:f eqn="sum @8 21600 0"/>
              <v:f eqn="prod @7 21600 pixelHeight"/>
              <v:f eqn="sum @10 21600 0"/>
            </v:formulas>
            <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
            <o:lock v:ext="edit" aspectratio="t"/>
          </v:shapetype>
          <v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:{width}pt;height:{height}pt" o:ole="">
            <v:imagedata r:id="rIdPreview" o:title=""/>
          </v:shape>
          <o:OLEObject Type="Embed" ProgID="{PROG_ID}" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_{object_id}" r:id="rIdOle"/>
        </w:object>
      </w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>"##
    );
    build_stored_zip(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", package_relationships.as_bytes()),
        ("word/document.xml", document.as_bytes()),
        ("word/_rels/document.xml.rels", document_relationships.as_bytes()),
        ("word/media/formula.emf", emf),
        ("word/embeddings/oleObject1.bin", storage),
    ])
}

/// Returns the storages of the VisualTeX OLE objects in a Flat OPC package
/// (`Range.WordOpenXML`), in document order.
pub fn visualtex_ole_storages(flat_opc: &str) -> Result<Vec<Vec<u8>>, String> {
    const DOCUMENT_PART: &str = "/word/document.xml";
    const DOCUMENT_RELATIONSHIPS_PART: &str = "/word/_rels/document.xml.rels";

    let mut reader = Reader::from_str(flat_opc);
    let mut part = String::new();
    let mut binary: Option<String> = None;
    let mut binaries: HashMap<String, String> = HashMap::new();
    let mut relationships: HashMap<String, String> = HashMap::new();
    let mut objects: Vec<String> = Vec::new();
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match &event {
            Event::Start(element) | Event::Empty(element) => {
                let attribute = |name: &[u8]| -> Result<Option<String>, String> {
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(xml_error)?;
                        if attribute.key.local_name().as_ref() == name {
                            let value = attribute
                                .normalized_value(XmlVersion::Implicit1_0)
                                .map_err(xml_error)?;
                            return Ok(Some(value.into_owned()));
                        }
                    }
                    Ok(None)
                };
                match element.local_name().as_ref() {
                    b"part" => part = attribute(b"name")?.unwrap_or_default(),
                    b"binaryData" if matches!(event, Event::Start(_)) => binary = Some(String::new()),
                    b"Relationship" if part == DOCUMENT_RELATIONSHIPS_PART => {
                        if let (Some(id), Some(target)) = (attribute(b"Id")?, attribute(b"Target")?) {
                            relationships.insert(id, target);
                        }
                    }
                    b"OLEObject" if part == DOCUMENT_PART => {
                        if attribute(b"ProgID")?.as_deref() == Some(PROG_ID) {
                            // `r:id`; `ShapeID` and `ObjectID` have other local names.
                            let id = attribute(b"id")?
                                .ok_or_else(|| "A VisualTeX OLE object has no relationship id".to_string())?;
                            objects.push(id);
                        }
                    }
                    _ => {}
                }
            }
            Event::Text(text) => {
                if let Some(binary) = binary.as_mut() {
                    binary.push_str(&text.decode().map_err(xml_error)?);
                }
            }
            Event::End(element) => match element.local_name().as_ref() {
                b"binaryData" => {
                    if let Some(data) = binary.take() {
                        binaries.insert(part.clone(), data);
                    }
                }
                b"part" => part.clear(),
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
    }

    objects
        .iter()
        .map(|id| {
            let target = relationships
                .get(id)
                .ok_or_else(|| format!("The VisualTeX OLE relationship {id} is missing"))?;
            let name = if target.starts_with('/') {
                target.clone()
            } else {
                format!("/word/{target}")
            };
            let encoded: String = binaries
                .get(&name)
                .ok_or_else(|| format!("The VisualTeX OLE part {name} is missing"))?
                .chars()
                .filter(|character| !character.is_ascii_whitespace())
                .collect();
            BASE64_STANDARD
                .decode(encoded)
                .map_err(|error| format!("The VisualTeX OLE part {name} is not Base64: {error}"))
        })
        .collect()
}

/// `CompObjStream` (MS-OLEDS 2.3.8): header, ANSI user type, no native
/// clipboard format, ANSI ProgID, then the empty Unicode extension.
fn comp_obj_stream() -> Vec<u8> {
    let mut stream = Vec::new();
    stream.extend_from_slice(&0xFFFE_0001_u32.to_le_bytes());
    stream.extend_from_slice(&0x0000_0A03_u32.to_le_bytes());
    stream.extend_from_slice(&0xFFFF_FFFF_u32.to_le_bytes());
    stream.extend_from_slice(&CLSID.to_bytes_le());
    for value in [Some(USER_TYPE), None, Some(PROG_ID)] {
        match value {
            Some(text) => {
                stream.extend_from_slice(&(text.len() as u32 + 1).to_le_bytes());
                stream.extend_from_slice(text.as_bytes());
                stream.push(0);
            }
            None => stream.extend_from_slice(&0_u32.to_le_bytes()),
        }
    }
    stream.extend_from_slice(&0x71B2_39F4_u32.to_le_bytes());
    stream.extend_from_slice(&[0; 12]);
    stream
}

/// `OLEStream` (MS-OLEDS 2.3.3) for an embedded, non-linked object.
fn ole_stream() -> Vec<u8> {
    let mut stream = Vec::new();
    stream.extend_from_slice(&0x0200_0001_u32.to_le_bytes());
    stream.extend_from_slice(&[0; 16]);
    stream
}

/// Word's `ObjInfo` (MS-DOC 2.9.155): no flags, EMF presentation.
fn obj_info_stream() -> Vec<u8> {
    let mut stream = Vec::new();
    stream.extend_from_slice(&0_u16.to_le_bytes());
    stream.extend_from_slice(&CF_ENHMETAFILE.to_le_bytes());
    stream.extend_from_slice(&0_u16.to_le_bytes());
    stream
}

fn format_points(value: f64) -> String {
    let rounded = (value * 100.0).round() / 100.0;
    format!("{rounded}")
}

fn storage_error(error: std::io::Error) -> String {
    format!("Unable to write the VisualTeX OLE storage: {error}")
}

fn xml_error(error: impl std::fmt::Display) -> String {
    format!("Unable to read the Word OLE package: {error}")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::office::sessions::MetadataLine;

    fn metadata() -> VisualTeXFormulaMetadata {
        VisualTeXFormulaMetadata {
            schema: "visualtex-formula".to_string(),
            schema_version: 1,
            formula_id: "1b4e28ba-2fa1-4d2c-883f-0016d3cca427".to_string(),
            title: String::new(),
            latex: "E=mc^2".to_string(),
            lines: vec![MetadataLine { id: "line-1".to_string(), latex: "E=mc^2".to_string() }],
            code_format: "latex".to_string(),
            display_mode: "inline".to_string(),
            inline_image_math_style: None,
            numbered: false,
            render_width_px: Some(60.0),
            render_height_px: Some(20.0),
            font_size_pt: Some(12.0),
            formula_letter_font: None,
            formula_chinese_font: None,
            reference_width_pt: None,
            reference_height_pt: None,
            reference_baseline_pt: None,
            image_ink_center_y_ratio: None,
            created_with_version: "1.2.8".to_string(),
            updated_with_version: "1.2.8".to_string(),
            created_at: "2026-10-05T00:00:00Z".to_string(),
            updated_at: "2026-10-05T00:00:00Z".to_string(),
        }
    }

    fn formula() -> OleFormula {
        let emf = crate::office::svg_emf::svg_to_vector_emf(
            br#"<svg xmlns="http://www.w3.org/2000/svg" width="60" height="20" viewBox="0 0 60 20"><path d="M2 2H58V18Z"/></svg>"#,
            60.0,
            20.0,
        )
        .unwrap();
        OleFormula {
            metadata: ole_metadata(None, &metadata(), Some(15.0)).unwrap(),
            emf,
            png: [PNG_SIGNATURE, b"rest"].concat(),
        }
    }

    #[test]
    fn storage_round_trips_with_the_windows_class_and_streams() {
        let formula = formula();
        let storage = build_ole_storage(&formula).unwrap();
        assert_eq!(read_ole_storage(&storage).unwrap(), formula);

        let file = cfb::CompoundFile::open(Cursor::new(&storage)).unwrap();
        assert_eq!(file.root_entry().clsid().to_string(), "8ff7f5aa-0d60-48d5-adbd-65a64b4c827b");
        let mut names: Vec<String> =
            file.read_root_storage().map(|entry| entry.name().to_string()).collect();
        names.sort();
        assert_eq!(
            names,
            ["\u{1}CompObj", "\u{1}Ole", "\u{3}ObjInfo", METADATA_STREAM, EMF_STREAM, PNG_STREAM]
        );
    }

    #[test]
    fn comp_obj_names_the_visualtex_class() {
        let stream = comp_obj_stream();
        assert_eq!(&stream[12..28], &CLSID.to_bytes_le());
        let text = String::from_utf8_lossy(&stream);
        assert!(text.contains("VisualTeX Formula\0"));
        assert!(text.contains("VisualTeX.Formula.1\0"));
        assert!(stream.ends_with(&[0xF4, 0x39, 0xB2, 0x71, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]));
    }

    #[test]
    fn rejects_foreign_objects_and_incomplete_payloads() {
        let mut file = cfb::CompoundFile::create_with_version(cfb::Version::V3, Cursor::new(Vec::new())).unwrap();
        file.create_stream("/Equation Native").unwrap();
        let foreign = file.into_inner().into_inner();
        assert!(read_ole_storage(&foreign).unwrap_err().contains("not a VisualTeX formula"));

        let mut formula = formula();
        formula.png = b"not a png".to_vec();
        let storage = build_ole_storage(&formula).unwrap();
        assert!(read_ole_storage(&storage).unwrap_err().contains("preview"));
    }

    #[test]
    fn metadata_keeps_unknown_windows_fields() {
        let mut original = ole_metadata(None, &metadata(), None).unwrap();
        original.insert("wordInlineOleWidthPt".to_string(), Value::from(31.5));
        original.insert("latex".to_string(), Value::from("old"));
        let mut edited = metadata();
        edited.latex = "a+b".to_string();
        let merged = ole_metadata(Some(&original), &edited, Some(12.5)).unwrap();
        assert_eq!(merged["latex"], "a+b");
        assert_eq!(merged["wordInlineOleWidthPt"], 31.5);
        assert_eq!(merged["baseline"], 12.5);
        assert_eq!(session_metadata(&merged).unwrap().latex, "a+b");
    }

    /// Word's `Range.WordOpenXML` form of a DOCX package.
    fn flat_opc(docx: &[u8]) -> String {
        let mut archive = zip::ZipArchive::new(Cursor::new(docx)).unwrap();
        let mut xml = String::from(
            r#"<?xml version="1.0" standalone="yes"?><?mso-application progid="Word.Document"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">"#,
        );
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index).unwrap();
            let name = format!("/{}", entry.name());
            let mut contents = Vec::new();
            entry.read_to_end(&mut contents).unwrap();
            if name.ends_with(".xml") || name.ends_with(".rels") {
                let text = String::from_utf8(contents).unwrap();
                let body = text.split_once("?>").map_or(text.as_str(), |(_, body)| body);
                xml.push_str(&format!(r#"<pkg:part pkg:name="{name}" pkg:contentType="application/xml"><pkg:xmlData>{body}</pkg:xmlData></pkg:part>"#));
            } else {
                let encoded = BASE64_STANDARD.encode(&contents);
                let wrapped: Vec<&str> = encoded.as_bytes().chunks(76).map(|line| std::str::from_utf8(line).unwrap()).collect();
                xml.push_str(&format!(
                    r#"<pkg:part pkg:name="{name}" pkg:contentType="application/octet-stream" pkg:compression="store"><pkg:binaryData>{}</pkg:binaryData></pkg:part>"#,
                    wrapped.join("\n")
                ));
            }
        }
        xml.push_str("</pkg:package>");
        xml
    }

    #[test]
    fn staging_docx_round_trips_through_word_open_xml() {
        let formula = formula();
        let storage = build_ole_storage(&formula).unwrap();
        let docx = build_word_ole_docx(&storage, &formula.emf, 45.25, 15.0, 1_234_567).unwrap();

        let mut archive = zip::ZipArchive::new(Cursor::new(&docx)).unwrap();
        let mut document = String::new();
        archive.by_name("word/document.xml").unwrap().read_to_string(&mut document).unwrap();
        assert!(document.contains(r#"<w:object w:dxaOrig="905" w:dyaOrig="300">"#));
        assert!(document.contains("style=\"width:45.25pt;height:15pt\""));
        assert!(document.contains(r#"ProgID="VisualTeX.Formula.1""#));
        assert!(document.contains(r#"ObjectID="_1234567""#));

        let storages = visualtex_ole_storages(&flat_opc(&docx)).unwrap();
        assert_eq!(storages, vec![storage]);
        assert_eq!(read_ole_storage(&storages[0]).unwrap(), formula);
    }

    #[test]
    fn word_open_xml_ignores_other_ole_classes() {
        let docx = build_word_ole_docx(b"x", b"y", 10.0, 10.0, 1).unwrap();
        let xml = flat_opc(&docx).replace(PROG_ID, "Equation.DSMT4");
        assert!(visualtex_ole_storages(&xml).unwrap().is_empty());
    }
}
