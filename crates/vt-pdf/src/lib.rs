use std::fs::{self, File};
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};

use image::{DynamicImage, ImageFormat, RgbaImage, imageops};
use pdfium_render::prelude::{PdfPageRenderRotation, PdfRenderConfig, Pdfium};
use sha2::{Digest, Sha256};
use vt_protocol::{
    PdfDocumentInfo, PdfPageInfo, PdfPixelDiffPage, PdfPixelDiffReport, PdfPixelRect,
    PdfRenderRequest, PdfRenderedImage,
};

static TEMP_SEQUENCE: AtomicU64 = AtomicU64::new(0);

const MIN_RENDER_WIDTH: u32 = 96;
const MAX_RENDER_WIDTH: u32 = 8_192;

#[derive(Debug, thiserror::Error)]
pub enum PdfError {
    #[error("PDF file does not exist or is not a regular file: {0}")]
    InvalidPdf(PathBuf),
    #[error("unsupported PDF render width {0}; expected {MIN_RENDER_WIDTH}..={MAX_RENDER_WIDTH}")]
    InvalidRenderWidth(u32),
    #[error("PDF page index {index} is out of bounds for {page_count} pages")]
    PageOutOfBounds { index: u32, page_count: u32 },
    #[error("tile rectangle is outside the rendered page")]
    TileOutOfBounds,
    #[error("PDFium is unavailable: {0}")]
    PdfiumUnavailable(String),
    #[error("PDFium operation failed: {0}")]
    Pdfium(String),
    #[error("image operation failed: {0}")]
    Image(String),
    #[error(transparent)]
    Io(#[from] std::io::Error),
}

#[derive(Clone, Debug)]
pub struct PdfService {
    cache_root: PathBuf,
}

impl PdfService {
    pub fn new(cache_root: impl Into<PathBuf>) -> Self {
        Self {
            cache_root: cache_root.into(),
        }
    }

    pub fn cache_root(&self) -> &Path {
        &self.cache_root
    }

    pub fn document_info(&self, pdf_path: impl AsRef<Path>) -> Result<PdfDocumentInfo, PdfError> {
        let pdf_path = canonical_pdf(pdf_path.as_ref())?;
        let metadata = fs::metadata(&pdf_path)?;
        let fingerprint = fingerprint_file(&pdf_path)?;
        let pdfium = pdfium()?;
        let document = pdfium
            .load_pdf_from_file(&pdf_path, None)
            .map_err(pdfium_error)?;
        let mut pages = Vec::with_capacity(document.pages().len() as usize);

        for index in 0..document.pages().len() {
            let page = document.pages().get(index).map_err(pdfium_error)?;
            pages.push(PdfPageInfo {
                index: index as u32,
                width_points: page.width().value,
                height_points: page.height().value,
                rotation_degrees: rotation_degrees(page.rotation().map_err(pdfium_error)?),
            });
        }

        Ok(PdfDocumentInfo {
            pdf_path,
            fingerprint,
            byte_len: metadata.len(),
            pages,
        })
    }

    pub fn render(&self, request: &PdfRenderRequest) -> Result<PdfRenderedImage, PdfError> {
        if !(MIN_RENDER_WIDTH..=MAX_RENDER_WIDTH).contains(&request.target_width_pixels) {
            return Err(PdfError::InvalidRenderWidth(request.target_width_pixels));
        }

        let info = self.document_info(&request.pdf_path)?;
        let page_count = info.pages.len() as u32;
        if request.page_index >= page_count {
            return Err(PdfError::PageOutOfBounds {
                index: request.page_index,
                page_count,
            });
        }

        let cache_path = self.cache_path(&info.fingerprint, request);
        if cache_path.is_file() {
            let dimensions = image::image_dimensions(&cache_path)
                .map_err(|error| PdfError::Image(error.to_string()))?;
            let page_dimensions = rendered_page_dimensions(
                &info.pages[request.page_index as usize],
                request.target_width_pixels,
            );
            return Ok(PdfRenderedImage {
                pdf_fingerprint: info.fingerprint,
                page_index: request.page_index,
                page_width_pixels: page_dimensions.0,
                page_height_pixels: page_dimensions.1,
                image_width_pixels: dimensions.0,
                image_height_pixels: dimensions.1,
                tile: request.tile,
                cache_path: cache_path.canonicalize()?,
                cache_hit: true,
            });
        }

        let pdfium = pdfium()?;
        let document = pdfium
            .load_pdf_from_file(&info.pdf_path, None)
            .map_err(pdfium_error)?;
        let page = document
            .pages()
            .get(request.page_index as u16)
            .map_err(pdfium_error)?;
        let render_config = PdfRenderConfig::new()
            .set_target_width(request.target_width_pixels as i32)
            .use_grayscale_rendering(request.grayscale)
            .render_annotations(true)
            .render_form_data(true);
        let bitmap = page
            .render_with_config(&render_config)
            .map_err(pdfium_error)?;
        let rendered = bitmap.as_image().into_rgba8();
        let page_width_pixels = rendered.width();
        let page_height_pixels = rendered.height();
        let output = crop_tile(rendered, request.tile)?;
        let image_width_pixels = output.width();
        let image_height_pixels = output.height();

        save_png_atomic(DynamicImage::ImageRgba8(output), &cache_path)?;

        Ok(PdfRenderedImage {
            pdf_fingerprint: info.fingerprint,
            page_index: request.page_index,
            page_width_pixels,
            page_height_pixels,
            image_width_pixels,
            image_height_pixels,
            tile: request.tile,
            cache_path: cache_path.canonicalize()?,
            cache_hit: false,
        })
    }

    pub fn compare_documents(
        &self,
        left_pdf: impl AsRef<Path>,
        right_pdf: impl AsRef<Path>,
        target_width_pixels: u32,
        tolerance: u8,
    ) -> Result<PdfPixelDiffReport, PdfError> {
        if !(MIN_RENDER_WIDTH..=MAX_RENDER_WIDTH).contains(&target_width_pixels) {
            return Err(PdfError::InvalidRenderWidth(target_width_pixels));
        }
        let left = self.document_info(left_pdf)?;
        let right = self.document_info(right_pdf)?;
        let page_count_matches = left.pages.len() == right.pages.len();
        let comparable_pages = left.pages.len().min(right.pages.len());
        let mut pages = Vec::with_capacity(comparable_pages);

        for page_index in 0..comparable_pages {
            let left_rendered = self.render(&PdfRenderRequest {
                pdf_path: left.pdf_path.clone(),
                page_index: page_index as u32,
                target_width_pixels,
                tile: None,
                grayscale: false,
            })?;
            let right_rendered = self.render(&PdfRenderRequest {
                pdf_path: right.pdf_path.clone(),
                page_index: page_index as u32,
                target_width_pixels,
                tile: None,
                grayscale: false,
            })?;
            let left_image = image::open(&left_rendered.cache_path)
                .map_err(|error| PdfError::Image(error.to_string()))?
                .into_rgba8();
            let right_image = image::open(&right_rendered.cache_path)
                .map_err(|error| PdfError::Image(error.to_string()))?
                .into_rgba8();
            pages.push(compare_rgba_images(
                page_index as u32,
                &left_image,
                &right_image,
                tolerance,
            ));
        }

        let maximum_changed_ratio = pages
            .iter()
            .map(|page| page.changed_ratio)
            .fold(if page_count_matches { 0.0 } else { 1.0 }, f64::max);

        Ok(PdfPixelDiffReport {
            page_count_matches,
            left_page_count: left.pages.len() as u32,
            right_page_count: right.pages.len() as u32,
            tolerance,
            pages,
            maximum_changed_ratio,
        })
    }

    pub fn purge_except(&self, current_fingerprint: &str) -> Result<usize, PdfError> {
        if !self.cache_root.exists() {
            return Ok(0);
        }
        let mut removed = 0;
        for entry in fs::read_dir(&self.cache_root)? {
            let entry = entry?;
            if entry.file_type()?.is_dir()
                && entry.file_name().to_string_lossy() != current_fingerprint
            {
                fs::remove_dir_all(entry.path())?;
                removed += 1;
            }
        }
        Ok(removed)
    }

    fn cache_path(&self, fingerprint: &str, request: &PdfRenderRequest) -> PathBuf {
        let mode = if request.grayscale { "gray" } else { "color" };
        let file_name = match request.tile {
            Some(tile) => format!(
                "page-{:05}-w{}-{mode}-tile-{}-{}-{}-{}.png",
                request.page_index,
                request.target_width_pixels,
                tile.x,
                tile.y,
                tile.width,
                tile.height
            ),
            None => format!(
                "page-{:05}-w{}-{mode}.png",
                request.page_index, request.target_width_pixels
            ),
        };
        self.cache_root.join(fingerprint).join(file_name)
    }
}

fn pdfium() -> Result<Pdfium, PdfError> {
    pdfium_auto::bind_bundled().map_err(|error| PdfError::PdfiumUnavailable(error.to_string()))
}

fn canonical_pdf(path: &Path) -> Result<PathBuf, PdfError> {
    if !path.is_file() {
        return Err(PdfError::InvalidPdf(path.to_path_buf()));
    }
    let canonical = path.canonicalize()?;
    if canonical.extension().and_then(|value| value.to_str()) != Some("pdf") {
        return Err(PdfError::InvalidPdf(canonical));
    }
    Ok(canonical)
}

fn fingerprint_file(path: &Path) -> Result<String, PdfError> {
    let mut file = File::open(path)?;
    let mut digest = Sha256::new();
    let mut buffer = [0_u8; 64 * 1024];
    loop {
        let count = file.read(&mut buffer)?;
        if count == 0 {
            break;
        }
        digest.update(&buffer[..count]);
    }
    Ok(format!("{:x}", digest.finalize()))
}

fn rotation_degrees(rotation: PdfPageRenderRotation) -> i16 {
    match rotation {
        PdfPageRenderRotation::None => 0,
        PdfPageRenderRotation::Degrees90 => 90,
        PdfPageRenderRotation::Degrees180 => 180,
        PdfPageRenderRotation::Degrees270 => 270,
    }
}

fn rendered_page_dimensions(page: &PdfPageInfo, target_width: u32) -> (u32, u32) {
    let (width_points, height_points) = if matches!(page.rotation_degrees, 90 | 270) {
        (page.height_points, page.width_points)
    } else {
        (page.width_points, page.height_points)
    };
    let ratio = if width_points > 0.0 {
        height_points / width_points
    } else {
        1.0
    };
    (
        target_width,
        ((target_width as f32) * ratio).round().max(1.0) as u32,
    )
}

fn compare_rgba_images(
    page_index: u32,
    left: &RgbaImage,
    right: &RgbaImage,
    tolerance: u8,
) -> PdfPixelDiffPage {
    if left.dimensions() != right.dimensions() {
        let width = left.width().max(right.width());
        let height = left.height().max(right.height());
        let total_pixels = u64::from(width) * u64::from(height);
        return PdfPixelDiffPage {
            page_index,
            width_pixels: width,
            height_pixels: height,
            changed_pixels: total_pixels,
            total_pixels,
            changed_ratio: 1.0,
            maximum_channel_delta: u8::MAX,
            mean_absolute_channel_delta: f64::from(u8::MAX),
        };
    }

    let mut changed_pixels = 0_u64;
    let mut maximum_channel_delta = 0_u8;
    let mut channel_delta_sum = 0_u64;
    for (left_pixel, right_pixel) in left.pixels().zip(right.pixels()) {
        let mut pixel_changed = false;
        for channel in 0..4 {
            let delta = left_pixel[channel].abs_diff(right_pixel[channel]);
            maximum_channel_delta = maximum_channel_delta.max(delta);
            channel_delta_sum += u64::from(delta);
            pixel_changed |= delta > tolerance;
        }
        if pixel_changed {
            changed_pixels += 1;
        }
    }
    let total_pixels = u64::from(left.width()) * u64::from(left.height());
    let total_channels = total_pixels.saturating_mul(4);
    PdfPixelDiffPage {
        page_index,
        width_pixels: left.width(),
        height_pixels: left.height(),
        changed_pixels,
        total_pixels,
        changed_ratio: if total_pixels == 0 {
            0.0
        } else {
            changed_pixels as f64 / total_pixels as f64
        },
        maximum_channel_delta,
        mean_absolute_channel_delta: if total_channels == 0 {
            0.0
        } else {
            channel_delta_sum as f64 / total_channels as f64
        },
    }
}

fn crop_tile(image: RgbaImage, tile: Option<PdfPixelRect>) -> Result<RgbaImage, PdfError> {
    let Some(tile) = tile else {
        return Ok(image);
    };
    let right = tile
        .x
        .checked_add(tile.width)
        .ok_or(PdfError::TileOutOfBounds)?;
    let bottom = tile
        .y
        .checked_add(tile.height)
        .ok_or(PdfError::TileOutOfBounds)?;
    if tile.width == 0 || tile.height == 0 || right > image.width() || bottom > image.height() {
        return Err(PdfError::TileOutOfBounds);
    }
    Ok(imageops::crop_imm(&image, tile.x, tile.y, tile.width, tile.height).to_image())
}

fn save_png_atomic(image: DynamicImage, destination: &Path) -> Result<(), PdfError> {
    let parent = destination
        .parent()
        .ok_or_else(|| PdfError::InvalidPdf(destination.to_path_buf()))?;
    fs::create_dir_all(parent)?;
    let sequence = TEMP_SEQUENCE.fetch_add(1, Ordering::Relaxed);
    let file_name = destination
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("page.png");
    let temporary = parent.join(format!(
        ".{file_name}.{}.{}.tmp",
        std::process::id(),
        sequence
    ));

    image
        .save_with_format(&temporary, ImageFormat::Png)
        .map_err(|error| PdfError::Image(error.to_string()))?;
    File::open(&temporary)?.sync_all()?;
    match fs::rename(&temporary, destination) {
        Ok(()) => {}
        Err(error) if destination.is_file() => {
            let _ = fs::remove_file(&temporary);
            if !destination.is_file() {
                return Err(error.into());
            }
        }
        Err(error) => return Err(error.into()),
    }
    if let Ok(directory) = File::open(parent) {
        let _ = directory.sync_all();
    }
    Ok(())
}

fn pdfium_error(error: impl std::fmt::Display) -> PdfError {
    PdfError::Pdfium(error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rendered_dimensions_preserve_aspect_ratio_and_rotation() {
        let mut page = PdfPageInfo {
            index: 0,
            width_points: 595.0,
            height_points: 842.0,
            rotation_degrees: 0,
        };
        assert_eq!(rendered_page_dimensions(&page, 595), (595, 842));
        page.rotation_degrees = 90;
        assert_eq!(rendered_page_dimensions(&page, 842), (842, 595));
    }

    #[test]
    fn tile_validation_rejects_overflow_and_out_of_bounds() {
        let image = RgbaImage::new(100, 100);
        let request = PdfPixelRect {
            x: 90,
            y: 90,
            width: 20,
            height: 20,
        };
        assert!(matches!(
            crop_tile(image, Some(request)),
            Err(PdfError::TileOutOfBounds)
        ));
    }

    #[test]
    fn fingerprint_changes_with_contents() {
        let temp = tempfile::tempdir().unwrap();
        let path = temp.path().join("sample.pdf");
        fs::write(&path, b"first").unwrap();
        let first = fingerprint_file(&path).unwrap();
        fs::write(&path, b"second").unwrap();
        let second = fingerprint_file(&path).unwrap();
        assert_ne!(first, second);
    }

    #[test]
    fn pixel_diff_respects_tolerance() {
        let left = RgbaImage::from_pixel(2, 1, image::Rgba([10, 20, 30, 255]));
        let mut right = left.clone();
        right.put_pixel(1, 0, image::Rgba([12, 20, 30, 255]));
        let tolerated = compare_rgba_images(0, &left, &right, 2);
        assert_eq!(tolerated.changed_pixels, 0);
        assert_eq!(tolerated.maximum_channel_delta, 2);
        let strict = compare_rgba_images(0, &left, &right, 1);
        assert_eq!(strict.changed_pixels, 1);
        assert_eq!(strict.changed_ratio, 0.5);
    }

    #[test]
    fn pixel_diff_reports_dimension_mismatch() {
        let left = RgbaImage::new(2, 2);
        let right = RgbaImage::new(3, 2);
        let report = compare_rgba_images(4, &left, &right, 0);
        assert_eq!(report.page_index, 4);
        assert_eq!(report.changed_ratio, 1.0);
        assert_eq!(report.maximum_channel_delta, u8::MAX);
    }
}
