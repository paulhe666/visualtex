//! The files a commit hands to the add-ins: PNG and SVG results, the formula ink centre, and the Word packages built
//! as stored zip files.

use super::*;

pub(super) fn paeth_predictor(left: u8, up: u8, upper_left: u8) -> u8 {
    let left = i32::from(left);
    let up = i32::from(up);
    let upper_left = i32::from(upper_left);
    let estimate = left + up - upper_left;
    let left_distance = (estimate - left).abs();
    let up_distance = (estimate - up).abs();
    let upper_left_distance = (estimate - upper_left).abs();
    if left_distance <= up_distance && left_distance <= upper_left_distance {
        left as u8
    } else if up_distance <= upper_left_distance {
        up as u8
    } else {
        upper_left as u8
    }
}

pub(super) fn png_ink_center_y_ratio_from_bytes(bytes: &[u8]) -> Result<f64, String> {
    const SIGNATURE: &[u8; 8] = b"\x89PNG\r\n\x1a\n";
    if bytes.len() < SIGNATURE.len() || &bytes[..8] != SIGNATURE {
        return Err("Cached formula image is not a PNG file".to_string());
    }

    let mut offset = 8usize;
    let mut width = 0usize;
    let mut height = 0usize;
    let mut bit_depth = 0u8;
    let mut color_type = 0u8;
    let mut interlace = 0u8;
    let mut compressed = Vec::new();
    while offset + 12 <= bytes.len() {
        let length_bytes: [u8; 4] = bytes[offset..offset + 4]
            .try_into()
            .map_err(|_| "Cached formula PNG has a truncated chunk length".to_string())?;
        let length = u32::from_be_bytes(length_bytes) as usize;
        let chunk_start = offset + 8;
        let chunk_end = chunk_start
            .checked_add(length)
            .ok_or_else(|| "PNG chunk length overflowed".to_string())?;
        let crc_end = chunk_end
            .checked_add(4)
            .ok_or_else(|| "PNG chunk boundary overflowed".to_string())?;
        if crc_end > bytes.len() {
            return Err("Cached formula PNG has a truncated chunk".to_string());
        }
        let chunk_type = &bytes[offset + 4..offset + 8];
        match chunk_type {
            b"IHDR" => {
                if length != 13 {
                    return Err("Cached formula PNG has an invalid IHDR".to_string());
                }
                let width_bytes: [u8; 4] = bytes[chunk_start..chunk_start + 4]
                    .try_into()
                    .map_err(|_| "Cached formula PNG has a truncated width".to_string())?;
                let height_bytes: [u8; 4] = bytes[chunk_start + 4..chunk_start + 8]
                    .try_into()
                    .map_err(|_| "Cached formula PNG has a truncated height".to_string())?;
                width = u32::from_be_bytes(width_bytes) as usize;
                height = u32::from_be_bytes(height_bytes) as usize;
                bit_depth = bytes[chunk_start + 8];
                color_type = bytes[chunk_start + 9];
                if bytes[chunk_start + 10] != 0 || bytes[chunk_start + 11] != 0 {
                    return Err("Cached formula PNG uses unsupported compression/filter methods".to_string());
                }
                interlace = bytes[chunk_start + 12];
            }
            b"IDAT" => compressed.extend_from_slice(&bytes[chunk_start..chunk_end]),
            b"IEND" => break,
            _ => {}
        }
        offset = crc_end;
    }

    if width == 0 || height == 0 || width > 200_000 || height > 200_000 {
        return Err("Cached formula PNG has invalid dimensions".to_string());
    }
    if bit_depth != 8 || !matches!(color_type, 4 | 6) || interlace != 0 {
        return Err("Cached formula PNG is not a supported 8-bit alpha image".to_string());
    }
    let bytes_per_pixel = if color_type == 6 { 4usize } else { 2usize };
    let stride = width
        .checked_mul(bytes_per_pixel)
        .ok_or_else(|| "Cached formula PNG row is too large".to_string())?;
    let expected = height
        .checked_mul(stride + 1)
        .ok_or_else(|| "Cached formula PNG is too large".to_string())?;
    if expected > 512 * 1024 * 1024 {
        return Err("Cached formula PNG exceeds the decode limit".to_string());
    }

    let mut inflated = Vec::with_capacity(expected);
    ZlibDecoder::new(compressed.as_slice())
        .read_to_end(&mut inflated)
        .map_err(|error| format!("Unable to inflate cached formula PNG: {error}"))?;
    if inflated.len() != expected {
        return Err("Cached formula PNG has an unexpected decoded size".to_string());
    }

    let mut previous = vec![0u8; stride];
    let mut current = vec![0u8; stride];
    let mut minimum_y = None::<usize>;
    let mut maximum_y = 0usize;
    for y in 0..height {
        let row_start = y * (stride + 1);
        let filter = inflated[row_start];
        let source = &inflated[row_start + 1..row_start + 1 + stride];
        for index in 0..stride {
            let left = if index >= bytes_per_pixel {
                current[index - bytes_per_pixel]
            } else {
                0
            };
            let up = previous[index];
            let upper_left = if index >= bytes_per_pixel {
                previous[index - bytes_per_pixel]
            } else {
                0
            };
            current[index] = match filter {
                0 => source[index],
                1 => source[index].wrapping_add(left),
                2 => source[index].wrapping_add(up),
                3 => source[index].wrapping_add(((u16::from(left) + u16::from(up)) / 2) as u8),
                4 => source[index].wrapping_add(paeth_predictor(left, up, upper_left)),
                _ => return Err("Cached formula PNG uses an unsupported row filter".to_string()),
            };
        }
        let alpha_offset = bytes_per_pixel - 1;
        let row_has_ink = current
            .chunks_exact(bytes_per_pixel)
            .any(|pixel| pixel[alpha_offset] >= 16);
        if row_has_ink {
            minimum_y.get_or_insert(y);
            maximum_y = y;
        }
        std::mem::swap(&mut previous, &mut current);
        current.fill(0);
    }

    let minimum_y = minimum_y.ok_or_else(|| "Cached formula PNG contains no visible ink".to_string())?;
    let painted_center = (minimum_y + maximum_y + 1) as f64 / 2.0;
    let ratio = painted_center / height as f64;
    if !ratio.is_finite() || !(0.0..=1.0).contains(&ratio) {
        return Err("Cached formula PNG produced an invalid ink center".to_string());
    }
    Ok(ratio)
}

pub(super) fn cached_formula_ink_center_y_ratio(formula_id: &str) -> Result<f64, String> {
    let (_, _, png_path) = word_image_cache_paths(formula_id)?;
    let metadata = fs::symlink_metadata(&png_path)
        .map_err(|error| format!("Unable to inspect cached formula PNG {}: {error}", png_path.display()))?;
    if metadata.file_type().is_symlink() || !metadata.is_file() || metadata.len() == 0 || metadata.len() > 128 * 1024 * 1024 {
        return Err("Cached formula PNG has an invalid file type or size".to_string());
    }
    let bytes = fs::read(&png_path)
        .map_err(|error| format!("Unable to read cached formula PNG {}: {error}", png_path.display()))?;
    png_ink_center_y_ratio_from_bytes(&bytes)
}

pub fn run_image_ink_center_cli_if_requested() -> Option<i32> {
    let arguments = std::env::args().collect::<Vec<_>>();
    let Some(index) = arguments.iter().position(|argument| argument == IMAGE_INK_CENTER_CLI_ARGUMENT) else {
        return None;
    };
    let formula_id = arguments.get(index + 1).map(String::as_str).unwrap_or("");
    match cached_formula_ink_center_y_ratio(formula_id) {
        Ok(ratio) => {
            println!("{ratio:.9}");
            Some(0)
        }
        Err(error) => {
            eprintln!("{error}");
            Some(2)
        }
    }
}

pub(super) fn materialize_result_png(session: &OfficeFormulaSession) -> Result<PathBuf, String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Office Session has no formula export".to_string())?;
    let png = export
        .png_base64
        .as_deref()
        .ok_or_else(|| "Offline Office Session requires a PNG export".to_string())
        .and_then(decode_png)?;
    let path = result_png_path(session.host, &session.id)?;
    atomic_write_runtime(&path, &png, 0o600)?;
    Ok(path)
}

pub(super) fn decode_svg(value: &str) -> Result<Vec<u8>, String> {
    let bytes = BASE64_STANDARD
        .decode(value.trim())
        .map_err(|error| format!("Unable to decode the Office SVG export: {error}"))?;
    if bytes.is_empty() || bytes.len() > MAX_METADATA_BYTES * 4 {
        return Err("Office formula SVG export is empty or too large".to_string());
    }
    let svg = std::str::from_utf8(&bytes)
        .map_err(|_| "Office formula SVG export is not UTF-8".to_string())?;
    let normalized = svg.trim_start();
    if !normalized.starts_with("<svg")
        && !(normalized.starts_with("<?xml") && normalized.contains("<svg"))
    {
        return Err("Office formula export is not a valid SVG document".to_string());
    }
    let lower = normalized.to_ascii_lowercase();
    for forbidden in [
        "<!doctype",
        "<!entity",
        "<script",
        "<foreignobject",
        "href=\"http:",
        "href=\"https:",
        "href=\"//",
        "xlink:href=\"http:",
        "xlink:href=\"https:",
        "xlink:href=\"//",
    ] {
        if lower.contains(forbidden) {
            return Err("Office formula SVG export contains unsafe external content".to_string());
        }
    }
    Ok(bytes)
}

pub(super) fn materialize_result_svg(session: &OfficeFormulaSession) -> Result<PathBuf, String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Office Session has no formula export".to_string())?;
    let svg = decode_svg(&export.svg_base64)?;
    let path = result_svg_path(session.host, &session.id)?;
    atomic_write_runtime(&path, &svg, 0o600)?;
    Ok(path)
}

pub(super) fn decode_stabilized_word_svg(value: &str) -> Result<Vec<u8>, String> {
    let svg = decode_svg(value)?;
    crate::svg_font_stabilizer::stabilize_word_svg_font_outlines(&svg)
}

pub(super) fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xffff_ffff_u32;
    for byte in bytes {
        crc ^= u32::from(*byte);
        for _ in 0..8 {
            let mask = 0_u32.wrapping_sub(crc & 1);
            crc = (crc >> 1) ^ (0xedb8_8320 & mask);
        }
    }
    !crc
}

pub(super) fn push_zip_u16(output: &mut Vec<u8>, value: u16) {
    output.extend_from_slice(&value.to_le_bytes());
}

pub(super) fn push_zip_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

pub(super) struct StoredZipEntry {
    pub(super) name: Vec<u8>,
    pub(super) crc: u32,
    pub(super) size: u32,
    pub(super) offset: u32,
}

pub(in crate::office) fn build_stored_zip<N, C>(entries: &[(N, C)]) -> Result<Vec<u8>, String>
where
    N: AsRef<str>,
    C: AsRef<[u8]>,
{
    let entry_count = u16::try_from(entries.len())
        .map_err(|_| "Word SVG staging package has too many ZIP entries".to_string())?;
    let mut output = Vec::new();
    let mut records = Vec::with_capacity(entries.len());

    for (name, contents) in entries {
        let name = name.as_ref();
        let contents = contents.as_ref();
        let name_bytes = name.as_bytes();
        let name_length = u16::try_from(name_bytes.len())
            .map_err(|_| "Word SVG staging package contains an overlong ZIP path".to_string())?;
        let size = u32::try_from(contents.len())
            .map_err(|_| "Word SVG staging package entry is too large".to_string())?;
        let offset = u32::try_from(output.len())
            .map_err(|_| "Word SVG staging package is too large".to_string())?;
        let checksum = crc32(contents);

        push_zip_u32(&mut output, 0x0403_4b50);
        push_zip_u16(&mut output, 20);
        push_zip_u16(&mut output, 0x0800);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 33);
        push_zip_u32(&mut output, checksum);
        push_zip_u32(&mut output, size);
        push_zip_u32(&mut output, size);
        push_zip_u16(&mut output, name_length);
        push_zip_u16(&mut output, 0);
        output.extend_from_slice(name_bytes);
        output.extend_from_slice(contents);

        records.push(StoredZipEntry {
            name: name_bytes.to_vec(),
            crc: checksum,
            size,
            offset,
        });
    }

    let central_offset = u32::try_from(output.len())
        .map_err(|_| "Word SVG staging package is too large".to_string())?;
    for record in &records {
        let name_length = u16::try_from(record.name.len())
            .map_err(|_| "Word SVG staging package contains an overlong ZIP path".to_string())?;
        push_zip_u32(&mut output, 0x0201_4b50);
        push_zip_u16(&mut output, 20);
        push_zip_u16(&mut output, 20);
        push_zip_u16(&mut output, 0x0800);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 33);
        push_zip_u32(&mut output, record.crc);
        push_zip_u32(&mut output, record.size);
        push_zip_u32(&mut output, record.size);
        push_zip_u16(&mut output, name_length);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u32(&mut output, 0);
        push_zip_u32(&mut output, record.offset);
        output.extend_from_slice(&record.name);
    }
    let central_size = u32::try_from(output.len())
        .map_err(|_| "Word SVG staging package is too large".to_string())?
        .checked_sub(central_offset)
        .ok_or_else(|| "Word SVG staging package central directory is invalid".to_string())?;

    push_zip_u32(&mut output, 0x0605_4b50);
    push_zip_u16(&mut output, 0);
    push_zip_u16(&mut output, 0);
    push_zip_u16(&mut output, entry_count);
    push_zip_u16(&mut output, entry_count);
    push_zip_u32(&mut output, central_size);
    push_zip_u32(&mut output, central_offset);
    push_zip_u16(&mut output, 0);
    Ok(output)
}

pub(super) fn build_word_svg_docx(
    svg: &[u8],
    png: &[u8],
    width_points: f64,
    height_points: f64,
) -> Result<Vec<u8>, String> {
    let width_emu = (width_points * 12_700.0).round();
    let height_emu = (height_points * 12_700.0).round();
    if !width_emu.is_finite()
        || !height_emu.is_finite()
        || width_emu <= 0.0
        || height_emu <= 0.0
        || width_emu > i64::MAX as f64
        || height_emu > i64::MAX as f64
    {
        return Err("Word SVG staging package dimensions are invalid".to_string());
    }
    let width_emu = width_emu as i64;
    let height_emu = height_emu as i64;

    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let package_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
    let document_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPng" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.png"/>
  <Relationship Id="rIdSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.svg"/>
</Relationships>"#;
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="{width_emu}" cy="{height_emu}"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:docPr id="1" name="VisualTeX Formula" descr="VisualTeX SVG formula"/>
            <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr><pic:cNvPr id="0" name="formula.svg"/><pic:cNvPicPr/></pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="rIdPng" cstate="print">
                      <a:extLst>
                        <a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext>
                      </a:extLst>
                    </a:blip>
                    <a:stretch><a:fillRect/></a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:noFill/><a:ln><a:noFill/></a:ln>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>"#
    );

    build_stored_zip(&[
        ("[Content_Types].xml", content_types.as_slice()),
        ("_rels/.rels", package_relationships.as_slice()),
        ("word/document.xml", document.as_bytes()),
        (
            "word/_rels/document.xml.rels",
            document_relationships.as_slice(),
        ),
        ("word/media/formula.png", png),
        ("word/media/formula.svg", svg),
    ])
}

pub(super) struct WordSvgBatchEntry {
    pub(super) svg: Vec<u8>,
    pub(super) png: Vec<u8>,
    pub(super) width_points: f64,
    pub(super) height_points: f64,
}

pub(super) fn build_word_svg_batch_docx(entries: &[WordSvgBatchEntry]) -> Result<Vec<u8>, String> {
    if entries.is_empty() || entries.len() > 1000 {
        return Err("Word SVG redraw batch must contain 1 to 1000 drawings".to_string());
    }
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let package_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
    let mut document_relationships = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
    );
    let mut document = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\">\n  <w:body>\n",
    );
    let mut zip_entries = vec![
        ("[Content_Types].xml".to_string(), content_types.to_vec()),
        ("_rels/.rels".to_string(), package_relationships.to_vec()),
    ];

    for (index, entry) in entries.iter().enumerate() {
        let width_emu = (entry.width_points * 12_700.0).round();
        let height_emu = (entry.height_points * 12_700.0).round();
        if !width_emu.is_finite()
            || !height_emu.is_finite()
            || width_emu <= 0.0
            || height_emu <= 0.0
            || width_emu > i64::MAX as f64
            || height_emu > i64::MAX as f64
        {
            return Err("Word SVG redraw batch dimensions are invalid".to_string());
        }
        let item_number = index + 1;
        let png_rel = format!("rIdPng{item_number}");
        let svg_rel = format!("rIdSvg{item_number}");
        let png_name = format!("formula-{item_number}.png");
        let svg_name = format!("formula-{item_number}.svg");
        document_relationships.push_str(&format!(
            "  <Relationship Id=\"{png_rel}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/{png_name}\"/>\n  <Relationship Id=\"{svg_rel}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/{svg_name}\"/>\n",
        ));
        document.push_str(&format!(
            r#"    <w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{}" cy="{}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{}" name="VisualTeX Formula {}" descr="VisualTeX SVG formula"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="{}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{}" cstate="print"><a:extLst><a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}"><asvg:svgBlip r:embed="{}"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{}" cy="{}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
"#,
            width_emu as i64,
            height_emu as i64,
            item_number,
            item_number,
            svg_name,
            png_rel,
            svg_rel,
            width_emu as i64,
            height_emu as i64,
        ));
        zip_entries.push((
            format!("word/media/{png_name}"),
            entry.png.clone(),
        ));
        zip_entries.push((
            format!("word/media/{svg_name}"),
            entry.svg.clone(),
        ));
    }
    document_relationships.push_str("</Relationships>");
    document.push_str(
        "    <w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr>\n  </w:body>\n</w:document>",
    );
    zip_entries.push((
        "word/document.xml".to_string(),
        document.into_bytes(),
    ));
    zip_entries.push((
        "word/_rels/document.xml.rels".to_string(),
        document_relationships.into_bytes(),
    ));
    build_stored_zip(&zip_entries)
}

pub(super) fn build_word_omml_batch_docx(omml_fragments: &[String]) -> Result<Vec<u8>, String> {
    if omml_fragments.is_empty() || omml_fragments.len() > 1000 {
        return Err("Word OMML batch must contain 1 to 1000 equations".to_string());
    }
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    let package_relationships = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
    let mut document = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><w:body>\n",
    );
    for fragment in omml_fragments {
        let normalized = fragment.trim();
        if !normalized.starts_with("<m:oMath")
            || !normalized.ends_with("</m:oMath>")
            || normalized.contains("<!DOCTYPE")
            || normalized.contains("<!ENTITY")
        {
            return Err("Word OMML batch contains an invalid equation fragment".to_string());
        }
        document.push_str("<w:p>");
        document.push_str(normalized);
        document.push_str("</w:p>\n");
    }
    document.push_str("<w:sectPr/></w:body></w:document>");
    build_stored_zip(&[
        ("[Content_Types].xml", content_types.as_slice()),
        ("_rels/.rels", package_relationships.as_slice()),
        ("word/document.xml", document.as_bytes()),
    ])
}

pub(super) fn materialize_word_svg_package(
    session: &OfficeFormulaSession,
    geometry: WordGeometry,
) -> Result<(PathBuf, PathBuf, PathBuf), String> {
    if session.host != OfficeHost::Word {
        return Err("Word SVG package materialization requires a Word Session".to_string());
    }
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Word Session has no formula export".to_string())?;
    let svg = decode_stabilized_word_svg(&export.svg_base64)?;
    let png = export
        .png_base64
        .as_deref()
        .ok_or_else(|| "Word SVG staging requires a PNG compatibility preview".to_string())
        .and_then(decode_png)?;
    let svg_path = result_svg_path(OfficeHost::Word, &session.id)?;
    let png_path = result_png_path(OfficeHost::Word, &session.id)?;
    let document_path = result_word_svg_docx_path(&session.id)?;
    atomic_write_runtime(&svg_path, &svg, 0o600)?;
    atomic_write_runtime(&png_path, &png, 0o600)?;
    let package = build_word_svg_docx(&svg, &png, geometry.width, geometry.height)?;
    atomic_write_runtime(&document_path, &package, 0o600)?;

    // Persist a formula-scoped image package even when the current Word output
    // is OMML. A later OMML-to-image conversion can then stay entirely inside
    // Word instead of waking the renderer and rebuilding the same SVG.
    let (cached_svg_path, cached_document_path, cached_png_path) =
        word_image_cache_paths(&session.formula_id)?;
    atomic_write_runtime(&cached_svg_path, &svg, 0o600)?;
    atomic_write_runtime(&cached_document_path, &package, 0o600)?;
    atomic_write_runtime(&cached_png_path, &png, 0o600)?;
    Ok((svg_path, document_path, png_path))
}

pub(super) fn materialize_powerpoint_svg(session: &OfficeFormulaSession) -> Result<PathBuf, String> {
    if session.host != OfficeHost::Powerpoint {
        return Err("PowerPoint SVG materialization requires a PowerPoint Session".to_string());
    }
    materialize_result_svg(session)
}
