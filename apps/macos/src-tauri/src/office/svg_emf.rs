//! Converts a VisualTeX formula SVG into a vector-only EMF.
//!
//! The Windows `VisualTeX.Formula.1` server refuses an OLE preview that
//! contains raster records, and Word draws the object from this EMF on both
//! platforms. The input is the Word export after font outlining, so it consists
//! of filled and stroked paths (MathJax glyphs, rules, enclosures) and nested
//! viewports. Anything that a plain EMF cannot express exactly (gradients,
//! partial transparency, masks, filters, images, live text) is rejected instead
//! of being approximated.

use usvg::tiny_skia_path::{PathSegment, Point};
use usvg::{FillRule, Group, LineCap, LineJoin, Node, Paint, PaintOrder, Transform};

/// EMF logical units per CSS pixel. The fixed window/viewport ratio keeps
/// glyph outlines at 1/32 px precision in 32-bit records.
const LOGICAL_UNITS_PER_PIXEL: f64 = 32.0;
/// Reference device: a 96 DPI display, 1920 x 1440 pixels = 508 x 381 mm.
/// The pixel and millimetre sizes are exact multiples of 480 px = 127 mm.
const REFERENCE_DEVICE_PIXELS: (i32, i32) = (1920, 1440);
const REFERENCE_DEVICE_MILLIMETERS: (i32, i32) = (508, 381);
const HUNDREDTHS_OF_MM_PER_PIXEL: f64 = 2540.0 / 96.0;

const EMR_HEADER: u32 = 1;
const EMR_POLYBEZIERTO: u32 = 5;
const EMR_SETWINDOWEXTEX: u32 = 9;
const EMR_SETVIEWPORTEXTEX: u32 = 11;
const EMR_EOF: u32 = 14;
const EMR_SETMAPMODE: u32 = 17;
const EMR_SETPOLYFILLMODE: u32 = 19;
const EMR_MOVETOEX: u32 = 27;
const EMR_SAVEDC: u32 = 33;
const EMR_RESTOREDC: u32 = 34;
const EMR_SELECTOBJECT: u32 = 37;
const EMR_CREATEBRUSHINDIRECT: u32 = 39;
const EMR_DELETEOBJECT: u32 = 40;
const EMR_LINETO: u32 = 54;
const EMR_SETMITERLIMIT: u32 = 58;
const EMR_BEGINPATH: u32 = 59;
const EMR_ENDPATH: u32 = 60;
const EMR_CLOSEFIGURE: u32 = 61;
const EMR_FILLPATH: u32 = 62;
const EMR_STROKEPATH: u32 = 64;
const EMR_SELECTCLIPPATH: u32 = 67;
const EMR_EXTCREATEPEN: u32 = 95;

const MM_ANISOTROPIC: u32 = 8;
const ALTERNATE: u32 = 1;
const WINDING: u32 = 2;
const RGN_AND: u32 = 1;
const BS_SOLID: u32 = 0;
const PS_GEOMETRIC: u32 = 0x0001_0000;
const PS_ENDCAP_ROUND: u32 = 0x0000;
const PS_ENDCAP_SQUARE: u32 = 0x0100;
const PS_ENDCAP_FLAT: u32 = 0x0200;
const PS_JOIN_ROUND: u32 = 0x0000;
const PS_JOIN_BEVEL: u32 = 0x1000;
const PS_JOIN_MITER: u32 = 0x2000;
const NULL_BRUSH: u32 = 0x8000_0005;
const BLACK_PEN: u32 = 0x8000_0007;
const BRUSH_HANDLE: u32 = 1;
const PEN_HANDLE: u32 = 2;

/// Converts `svg` into an EMF whose frame is `width_px` x `height_px` CSS
/// pixels at 96 DPI, the same frame the Windows add-in records for Word.
pub fn svg_to_vector_emf(svg: &[u8], width_px: f64, height_px: f64) -> Result<Vec<u8>, String> {
    if !(width_px.is_finite() && height_px.is_finite() && width_px > 0.0 && height_px > 0.0) {
        return Err("EMF preview dimensions must be positive".to_string());
    }
    let text = std::str::from_utf8(svg).map_err(|_| "Formula SVG is not UTF-8".to_string())?;
    let document = usvg::roxmltree::Document::parse(text)
        .map_err(|error| format!("Formula SVG is not valid XML: {error}"))?;
    for element in document.descendants().filter(|node| node.is_element()) {
        match element.tag_name().name() {
            // Live text would render with whatever font the reader has; the
            // Word export outlines every glyph before it reaches this point.
            "text" => return Err("Formula SVG still contains text; outline it before EMF conversion".to_string()),
            "image" | "foreignObject" => {
                return Err(format!(
                    "Formula SVG element <{}> cannot be stored in a vector EMF",
                    element.tag_name().name()
                ))
            }
            _ => {}
        }
    }
    let tree = usvg::Tree::from_xmltree(&document, &usvg::Options::default())
        .map_err(|error| format!("Unable to read the formula SVG: {error}"))?;

    let size = tree.size();
    let output = Transform::from_scale(
        (width_px * LOGICAL_UNITS_PER_PIXEL / f64::from(size.width())) as f32,
        (height_px * LOGICAL_UNITS_PER_PIXEL / f64::from(size.height())) as f32,
    );
    let mut writer = EmfWriter::new();
    writer.write_group(tree.root(), output, 1.0)?;
    Ok(writer.finish(width_px, height_px))
}

struct EmfWriter {
    body: Vec<u8>,
    records: u32,
    fill_rule: Option<u32>,
}

impl EmfWriter {
    fn new() -> Self {
        let mut writer = Self { body: Vec::new(), records: 0, fill_rule: None };
        writer.record(EMR_SETMAPMODE, &[MM_ANISOTROPIC]);
        let scale = 1000 * LOGICAL_UNITS_PER_PIXEL as u32;
        writer.record(EMR_SETWINDOWEXTEX, &[scale, scale]);
        writer.record(EMR_SETVIEWPORTEXTEX, &[1000, 1000]);
        writer
    }

    fn record(&mut self, kind: u32, fields: &[u32]) {
        self.body.extend_from_slice(&kind.to_le_bytes());
        self.body.extend_from_slice(&(8 + 4 * fields.len() as u32).to_le_bytes());
        for field in fields {
            self.body.extend_from_slice(&field.to_le_bytes());
        }
        self.records += 1;
    }

    fn write_group(&mut self, group: &Group, output: Transform, opacity: f32) -> Result<(), String> {
        if group.mask().is_some() || !group.filters().is_empty() {
            return Err("Formula SVG masks and filters cannot be stored in a vector EMF".to_string());
        }
        if group.blend_mode() != usvg::BlendMode::Normal {
            return Err("Formula SVG blend modes cannot be stored in a vector EMF".to_string());
        }
        let opacity = opacity * group.opacity().get();
        if let Some(clip) = group.clip_path() {
            if clip.clip_path().is_some() {
                return Err("Nested formula SVG clip paths are not supported".to_string());
            }
            self.record(EMR_SAVEDC, &[]);
            self.begin_path();
            let clip_transform = output.pre_concat(group.abs_transform()).pre_concat(clip.transform());
            let rule = self.append_clip_group(clip.root(), clip_transform)?;
            self.end_path();
            self.set_fill_rule(rule);
            self.record(EMR_SELECTCLIPPATH, &[RGN_AND]);
        }
        for child in group.children() {
            match child {
                Node::Group(child) => self.write_group(child, output, opacity)?,
                Node::Path(path) => self.write_path(path, output, opacity)?,
                Node::Image(_) | Node::Text(_) => {
                    return Err("Formula SVG images and text cannot be stored in a vector EMF".to_string())
                }
            }
        }
        if group.clip_path().is_some() {
            self.record(EMR_RESTOREDC, &[(-1_i32) as u32]);
            // RestoreDC also restores the fill mode selected before the clip.
            self.fill_rule = None;
        }
        Ok(())
    }

    /// Clip-path content uses transforms relative to the clip root, not the
    /// document, so accumulate them explicitly. Returns the clip fill rule.
    fn append_clip_group(&mut self, group: &Group, transform: Transform) -> Result<u32, String> {
        let mut rule = WINDING;
        for child in group.children() {
            match child {
                Node::Group(child) => {
                    if child.clip_path().is_some() {
                        return Err("Nested formula SVG clip paths are not supported".to_string());
                    }
                    rule = self.append_clip_group(child, transform.pre_concat(child.transform()))?;
                }
                Node::Path(path) => {
                    self.append_segments(path.data(), transform);
                    if let Some(fill) = path.fill() {
                        rule = fill_mode(fill.rule());
                    }
                }
                Node::Image(_) | Node::Text(_) => {
                    return Err("Formula SVG clip paths may contain only shapes".to_string())
                }
            }
        }
        Ok(rule)
    }

    fn write_path(&mut self, path: &usvg::Path, output: Transform, opacity: f32) -> Result<(), String> {
        if !path.is_visible() {
            return Ok(());
        }
        let transform = output.pre_concat(path.abs_transform());
        let fill = match path.fill() {
            Some(fill) => solid_color(fill.paint(), fill.opacity().get() * opacity)?
                .map(|color| (color, fill_mode(fill.rule()))),
            None => None,
        };
        let stroke = match path.stroke() {
            Some(stroke) => {
                if stroke.dasharray().is_some() {
                    return Err("Dashed formula SVG strokes cannot be stored in a vector EMF".to_string());
                }
                solid_color(stroke.paint(), stroke.opacity().get() * opacity)?.map(|color| (color, stroke))
            }
            None => None,
        };
        let stroke_first = path.paint_order() == PaintOrder::StrokeAndFill;
        if stroke_first {
            if let Some((color, stroke)) = stroke {
                self.stroke(path, transform, color, stroke);
            }
        }
        if let Some((color, rule)) = fill {
            self.fill(path, transform, color, rule);
        }
        if !stroke_first {
            if let Some((color, stroke)) = stroke {
                self.stroke(path, transform, color, stroke);
            }
        }
        Ok(())
    }

    fn fill(&mut self, path: &usvg::Path, transform: Transform, color: u32, rule: u32) {
        self.record(EMR_CREATEBRUSHINDIRECT, &[BRUSH_HANDLE, BS_SOLID, color, 0]);
        self.record(EMR_SELECTOBJECT, &[BRUSH_HANDLE]);
        self.set_fill_rule(rule);
        self.begin_path();
        self.append_segments(path.data(), transform);
        self.end_path();
        let bounds = device_bounds(path.data(), transform);
        self.record(EMR_FILLPATH, &bounds);
        self.record(EMR_SELECTOBJECT, &[NULL_BRUSH]);
        self.record(EMR_DELETEOBJECT, &[BRUSH_HANDLE]);
    }

    fn stroke(&mut self, path: &usvg::Path, transform: Transform, color: u32, stroke: &usvg::Stroke) {
        // A geometric pen width is in logical units and follows the mapping
        // mode. Non-uniform scaling uses the area-preserving average.
        let scale = f64::from((transform.sx * transform.sy - transform.kx * transform.ky).abs()).sqrt();
        let width = (f64::from(stroke.width().get()) * scale).round().max(1.0) as u32;
        let cap = match stroke.linecap() {
            LineCap::Butt => PS_ENDCAP_FLAT,
            LineCap::Round => PS_ENDCAP_ROUND,
            LineCap::Square => PS_ENDCAP_SQUARE,
        };
        let join = match stroke.linejoin() {
            LineJoin::Miter | LineJoin::MiterClip => PS_JOIN_MITER,
            LineJoin::Round => PS_JOIN_ROUND,
            LineJoin::Bevel => PS_JOIN_BEVEL,
        };
        self.record(
            EMR_EXTCREATEPEN,
            &[PEN_HANDLE, 0, 0, 0, 0, PS_GEOMETRIC | cap | join, width, BS_SOLID, color, 0, 0],
        );
        self.record(EMR_SELECTOBJECT, &[PEN_HANDLE]);
        self.record(EMR_SETMITERLIMIT, &[stroke.miterlimit().get().round().max(1.0) as u32]);
        self.begin_path();
        self.append_segments(path.data(), transform);
        self.end_path();
        let bounds = device_bounds(path.data(), transform);
        self.record(EMR_STROKEPATH, &bounds);
        self.record(EMR_SELECTOBJECT, &[BLACK_PEN]);
        self.record(EMR_DELETEOBJECT, &[PEN_HANDLE]);
    }

    fn set_fill_rule(&mut self, rule: u32) {
        if self.fill_rule != Some(rule) {
            self.record(EMR_SETPOLYFILLMODE, &[rule]);
            self.fill_rule = Some(rule);
        }
    }

    fn begin_path(&mut self) {
        self.record(EMR_BEGINPATH, &[]);
    }

    fn end_path(&mut self) {
        self.record(EMR_ENDPATH, &[]);
    }

    fn append_segments(&mut self, data: &usvg::tiny_skia_path::Path, transform: Transform) {
        let mut current = Point::zero();
        for segment in data.segments() {
            match segment {
                PathSegment::MoveTo(point) => {
                    current = point;
                    let (x, y) = logical(point, transform);
                    self.record(EMR_MOVETOEX, &[x as u32, y as u32]);
                }
                PathSegment::LineTo(point) => {
                    current = point;
                    let (x, y) = logical(point, transform);
                    self.record(EMR_LINETO, &[x as u32, y as u32]);
                }
                PathSegment::QuadTo(control, end) => {
                    // Degree elevation: exact cubic form of the quadratic.
                    let first = lerp(current, control, 2.0 / 3.0);
                    let second = lerp(end, control, 2.0 / 3.0);
                    self.bezier(&[first, second, end], transform);
                    current = end;
                }
                PathSegment::CubicTo(first, second, end) => {
                    self.bezier(&[first, second, end], transform);
                    current = end;
                }
                PathSegment::Close => self.record(EMR_CLOSEFIGURE, &[]),
            }
        }
    }

    fn bezier(&mut self, points: &[Point; 3], transform: Transform) {
        let logical_points = points.map(|point| logical(point, transform));
        let mut fields = Vec::with_capacity(5 + 6);
        let (mut left, mut top, mut right, mut bottom) = (i32::MAX, i32::MAX, i32::MIN, i32::MIN);
        for (x, y) in logical_points {
            left = left.min(device_floor(x));
            top = top.min(device_floor(y));
            right = right.max(device_ceil(x));
            bottom = bottom.max(device_ceil(y));
        }
        fields.extend([left as u32, top as u32, right as u32, bottom as u32, 3]);
        for (x, y) in logical_points {
            fields.extend([x as u32, y as u32]);
        }
        self.record(EMR_POLYBEZIERTO, &fields);
    }

    fn finish(mut self, width_px: f64, height_px: f64) -> Vec<u8> {
        self.record(EMR_EOF, &[0, 16, 20]);
        let header_size = 108_u32;
        let total = header_size + self.body.len() as u32;
        let device_right = width_px.ceil() as i32 - 1;
        let device_bottom = height_px.ceil() as i32 - 1;
        let frame_right = (width_px * HUNDREDTHS_OF_MM_PER_PIXEL).round() as i32;
        let frame_bottom = (height_px * HUNDREDTHS_OF_MM_PER_PIXEL).round() as i32;
        let fields: [u32; 25] = [
            0,
            0,
            device_right as u32,
            device_bottom as u32,
            0,
            0,
            frame_right as u32,
            frame_bottom as u32,
            0x464D_4520, // " EMF"
            0x0001_0000,
            total,
            self.records + 1,
            3, // handle table: index 0 plus the brush and pen slots
            0, // no description
            0,
            0, // no palette
            REFERENCE_DEVICE_PIXELS.0 as u32,
            REFERENCE_DEVICE_PIXELS.1 as u32,
            REFERENCE_DEVICE_MILLIMETERS.0 as u32,
            REFERENCE_DEVICE_MILLIMETERS.1 as u32,
            0, // no pixel format
            0,
            0, // not OpenGL
            REFERENCE_DEVICE_MILLIMETERS.0 as u32 * 1000,
            REFERENCE_DEVICE_MILLIMETERS.1 as u32 * 1000,
        ];
        let mut output = Vec::with_capacity(total as usize);
        output.extend_from_slice(&EMR_HEADER.to_le_bytes());
        output.extend_from_slice(&header_size.to_le_bytes());
        for (index, field) in fields.iter().enumerate() {
            // nHandles and sReserved share one 32-bit slot.
            if index == 12 {
                output.extend_from_slice(&(*field as u16).to_le_bytes());
                output.extend_from_slice(&0_u16.to_le_bytes());
            } else {
                output.extend_from_slice(&field.to_le_bytes());
            }
        }
        output.extend_from_slice(&self.body);
        output
    }
}

fn fill_mode(rule: FillRule) -> u32 {
    match rule {
        FillRule::NonZero => WINDING,
        FillRule::EvenOdd => ALTERNATE,
    }
}

/// Returns the COLORREF for a solid paint, `None` for an invisible one.
fn solid_color(paint: &Paint, opacity: f32) -> Result<Option<u32>, String> {
    let Paint::Color(color) = paint else {
        return Err("Formula SVG gradients and patterns cannot be stored in a vector EMF".to_string());
    };
    // MathJax paints a 0.001-opacity background rectangle for hit testing.
    if opacity * 255.0 < 1.0 {
        return Ok(None);
    }
    if opacity < 1.0 {
        return Err("Partially transparent formula SVG paint cannot be stored in a vector EMF".to_string());
    }
    Ok(Some(
        u32::from(color.red) | u32::from(color.green) << 8 | u32::from(color.blue) << 16,
    ))
}

fn lerp(from: Point, to: Point, amount: f32) -> Point {
    Point::from_xy(from.x + (to.x - from.x) * amount, from.y + (to.y - from.y) * amount)
}

fn logical(point: Point, transform: Transform) -> (i32, i32) {
    let mut mapped = point;
    transform.map_point(&mut mapped);
    (mapped.x.round() as i32, mapped.y.round() as i32)
}

fn device_floor(logical: i32) -> i32 {
    (f64::from(logical) / LOGICAL_UNITS_PER_PIXEL).floor() as i32
}

fn device_ceil(logical: i32) -> i32 {
    (f64::from(logical) / LOGICAL_UNITS_PER_PIXEL).ceil() as i32
}

fn device_bounds(data: &usvg::tiny_skia_path::Path, transform: Transform) -> [u32; 4] {
    let bounds = data.bounds();
    let corners = [
        Point::from_xy(bounds.left(), bounds.top()),
        Point::from_xy(bounds.right(), bounds.top()),
        Point::from_xy(bounds.left(), bounds.bottom()),
        Point::from_xy(bounds.right(), bounds.bottom()),
    ];
    let (mut left, mut top, mut right, mut bottom) = (i32::MAX, i32::MAX, i32::MIN, i32::MIN);
    for corner in corners {
        let (x, y) = logical(corner, transform);
        left = left.min(device_floor(x));
        top = top.min(device_floor(y));
        right = right.max(device_ceil(x));
        bottom = bottom.max(device_ceil(y));
    }
    [left as u32, top as u32, right as u32, bottom as u32]
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;

    /// Mirrors `IsVectorEmf` in the Windows LocalServer
    /// (apps/windows/src-windows/VisualTeX.FormulaOleServer/FormulaOleObject.cpp):
    /// a valid record chain, at least one vector record and no raster record.
    pub(crate) fn assert_windows_accepts_emf(bytes: &[u8]) {
        let read = |offset: usize| u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap());
        assert_eq!(read(0), EMR_HEADER);
        assert_eq!(read(40), 0x464D_4520);
        assert_eq!(read(48) as usize, bytes.len(), "nBytes");
        let raster = [76, 77, 78, 80, 81, 82, 92, 93, 114, 116];
        let vector = [2, 3, 4, 5, 6, 7, 8, 42, 43, 44, 45, 46, 47, 54, 55, 56, 62, 63, 64, 83, 84, 85, 86, 87, 88, 89, 90, 91, 96];
        let mut offset = 0;
        let mut records = 0;
        let mut saw_vector = false;
        let mut last = 0;
        while offset < bytes.len() {
            let kind = read(offset);
            let size = read(offset + 4) as usize;
            assert!(size >= 8 && size % 4 == 0 && offset + size <= bytes.len(), "record {kind} size {size}");
            assert!(!raster.contains(&kind), "raster record {kind}");
            saw_vector |= vector.contains(&kind);
            records += 1;
            last = kind;
            offset += size;
        }
        assert_eq!(offset, bytes.len());
        assert_eq!(last, EMR_EOF);
        assert_eq!(read(52), records, "nRecords");
        assert!(saw_vector, "no vector record");
    }

    fn svg(body: &str) -> Vec<u8> {
        format!(
            r#"<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="20" height="10" viewBox="0 0 2000 1000">{body}</svg>"#
        )
        .into_bytes()
    }

    #[test]
    fn writes_a_vector_emf_with_a_96_dpi_frame() {
        let emf = svg_to_vector_emf(&svg(r##"<path d="M0 0H1000V500Z" fill="#102030"/>"##), 20.0, 10.0).unwrap();
        assert_windows_accepts_emf(&emf);
        let read = |offset: usize| i32::from_le_bytes(emf[offset..offset + 4].try_into().unwrap());
        assert_eq!((read(24), read(28), read(32), read(36)), (0, 0, 529, 265));
        // The brush carries the exact COLORREF (0x00BBGGRR).
        let brush = emf.windows(4).position(|window| window == 0x0030_2010_u32.to_le_bytes()).unwrap();
        assert_eq!(read(brush - 16) as u32, EMR_CREATEBRUSHINDIRECT);
    }

    #[test]
    fn skips_mathjax_hit_area_and_rejects_partial_transparency() {
        let emf = svg_to_vector_emf(
            &svg(r##"<rect width="2000" height="1000" fill="#000" fill-opacity="0.001"/><path d="M0 0H10V10Z"/>"##),
            20.0,
            10.0,
        )
        .unwrap();
        assert_windows_accepts_emf(&emf);
        assert_eq!(emf.windows(4).filter(|w| *w == EMR_FILLPATH.to_le_bytes()).count(), 1);
        let error = svg_to_vector_emf(&svg(r#"<path d="M0 0H10V10Z" fill-opacity="0.5"/>"#), 20.0, 10.0).unwrap_err();
        assert!(error.contains("transparent"), "{error}");
    }

    #[test]
    fn rejects_text_images_and_gradients() {
        for body in [
            r#"<text x="0" y="10">x</text>"#,
            r#"<image width="1" height="1" href="data:image/png;base64,AAAA"/>"#,
            r##"<linearGradient id="g"><stop offset="0" stop-color="#000"/><stop offset="1" stop-color="#fff"/></linearGradient><path d="M0 0H10V10Z" fill="url(#g)"/>"##,
        ] {
            assert!(svg_to_vector_emf(&svg(body), 20.0, 10.0).is_err(), "{body}");
        }
    }

    #[test]
    fn nested_viewports_become_clip_paths() {
        let emf = svg_to_vector_emf(
            &svg(r#"<svg x="0" y="0" width="500" height="500" viewBox="0 0 500 500"><path d="M0 0H900V900Z"/></svg>"#),
            20.0,
            10.0,
        )
        .unwrap();
        assert_windows_accepts_emf(&emf);
        let kinds: Vec<u32> = record_kinds(&emf);
        let save = kinds.iter().position(|kind| *kind == EMR_SAVEDC).unwrap();
        let clip = kinds.iter().position(|kind| *kind == EMR_SELECTCLIPPATH).unwrap();
        let fill = kinds.iter().position(|kind| *kind == EMR_FILLPATH).unwrap();
        let restore = kinds.iter().position(|kind| *kind == EMR_RESTOREDC).unwrap();
        assert!(save < clip && clip < fill && fill < restore);
    }

    pub(crate) fn record_kinds(emf: &[u8]) -> Vec<u32> {
        let mut kinds = Vec::new();
        let mut offset = 0;
        while offset < emf.len() {
            kinds.push(u32::from_le_bytes(emf[offset..offset + 4].try_into().unwrap()));
            offset += u32::from_le_bytes(emf[offset + 4..offset + 8].try_into().unwrap()) as usize;
        }
        kinds
    }
}
