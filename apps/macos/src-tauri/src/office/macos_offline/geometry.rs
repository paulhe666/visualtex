//! Formula picture sizes in Word and PowerPoint: reference geometry, SVG geometry per font size, PowerPoint scaling.

use super::*;

pub(super) fn word_image_visual_scales(formula_letter_font: Option<&str>) -> (f64, f64) {
    match formula_letter_font {
        Some("times") => (WORD_TIMES_IMAGE_WIDTH_SCALE, WORD_TIMES_IMAGE_HEIGHT_SCALE),
        _ => (WORD_TEX_IMAGE_VISUAL_SCALE, WORD_TEX_IMAGE_VISUAL_SCALE),
    }
}

pub(super) fn scale_word_reference_geometry(
    reference_width_pt: f64,
    reference_height_pt: f64,
    reference_baseline_pt: f64,
    font_size_pt: f64,
) -> Result<WordGeometry, String> {
    if !reference_width_pt.is_finite()
        || !reference_height_pt.is_finite()
        || reference_width_pt <= 0.0
        || reference_height_pt <= 0.0
        || !reference_baseline_pt.is_finite()
        || !(-256.0..=0.0).contains(&reference_baseline_pt)
        || !font_size_pt.is_finite()
        || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&font_size_pt)
    {
        return Err("Word formula point-size reference geometry is invalid".to_string());
    }
    let point_scale = font_size_pt / WORD_REFERENCE_FONT_SIZE_PT;
    let width = reference_width_pt * point_scale;
    let height = reference_height_pt * point_scale;
    if !width.is_finite()
        || !height.is_finite()
        || width <= 0.0
        || height <= 0.0
        || width > 10_000.0
        || height > 10_000.0
    {
        return Err("Word formula point-size geometry is invalid".to_string());
    }
    // Word exposes Font.Position only as whole points. SVG and imported
    // picture dimensions are quantized independently, so a mathematical half
    // point can arrive just below the boundary (the real Times fixture reports
    // -1.491 pt instead of -1.5 pt). Use a small rounding tolerance for negative
    // descents; unlike a fixed offset, this changes only boundary cases and
    // preserves the distinct fraction/integral/sum/root baseline classes.
    let raw_baseline = reference_baseline_pt * point_scale;
    let baseline = if raw_baseline < 0.0 {
        (raw_baseline - 0.01).round()
    } else {
        0.0
    }
    .clamp(-256.0, 0.0) as i32;
    Ok(WordGeometry {
        width,
        height,
        baseline,
        font_size_pt,
        reference_width_pt,
        reference_height_pt,
        reference_baseline_pt,
    })
}

pub(super) fn calculate_word_svg_geometry_with_scale(
    width: f64,
    height: f64,
    baseline: f64,
    font_size_pt: f64,
    width_scale: f64,
    height_scale: f64,
) -> Result<WordGeometry, String> {
    if !width.is_finite()
        || !height.is_finite()
        || !baseline.is_finite()
        || width <= 0.0
        || height <= 0.0
        || baseline < 0.0
        || baseline > height
        || !width_scale.is_finite()
        || !(0.5..=2.0).contains(&width_scale)
        || !height_scale.is_finite()
        || !(0.5..=2.0).contains(&height_scale)
    {
        return Err("Word formula SVG geometry is invalid".to_string());
    }
    let natural_width = width * 0.75 * width_scale;
    let natural_height = height * 0.75 * height_scale;
    let reference_scale = f64::min(1.0, MAX_WORD_WIDTH_PT / natural_width);
    let reference_width_pt = natural_width * reference_scale;
    let reference_height_pt = natural_height * reference_scale;
    let descent_ratio = (height - baseline) / height;
    // Preserve the fractional descent at the canonical 14 pt size. Word's
    // Font.Position is integral, so rounding this reference and rounding again
    // after point-size scaling visibly under-corrects short subscript formulas
    // such as L_z beside L^2.
    let reference_baseline_pt = (-(reference_height_pt * descent_ratio).max(0.0))
        .clamp(-256.0, 0.0);
    scale_word_reference_geometry(
        reference_width_pt,
        reference_height_pt,
        reference_baseline_pt,
        font_size_pt,
    )
}

#[cfg(test)]
pub(super) fn calculate_word_svg_geometry(
    width: f64,
    height: f64,
    baseline: f64,
    font_size_pt: f64,
) -> Result<WordGeometry, String> {
    calculate_word_svg_geometry_with_scale(
        width,
        height,
        baseline,
        font_size_pt,
        WORD_TEX_IMAGE_VISUAL_SCALE,
        WORD_TEX_IMAGE_VISUAL_SCALE,
    )
}

pub(super) fn calculate_word_svg_geometry_for_font(
    width: f64,
    height: f64,
    baseline: f64,
    font_size_pt: f64,
    formula_letter_font: Option<&str>,
) -> Result<WordGeometry, String> {
    let (width_scale, height_scale) = word_image_visual_scales(formula_letter_font);
    let geometry = calculate_word_svg_geometry_with_scale(
        width,
        height,
        baseline,
        font_size_pt,
        width_scale,
        height_scale,
    )?;
    // The frontend aligns the SVG canvas to Word whole-point positions.
    // Preserve its mathematical baseline; a formula-specific descent floor
    // would shift only shallow formulas away from the shared text baseline.
    Ok(geometry)
}

pub(super) fn calculate_word_geometry(
    request: &MacOfflineSessionRequest,
    session: &OfficeFormulaSession,
) -> Result<WordGeometry, String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Word Session has no formula export".to_string())?;
    if !export.width.is_finite()
        || !export.height.is_finite()
        || export.width <= 0.0
        || export.height <= 0.0
    {
        return Err("Word formula export has invalid dimensions".to_string());
    }

    // The editor persists the user's current selection on the Session. That
    // value must win over the immutable launch request; otherwise changing the
    // size in the editor is silently replaced by the size observed when the
    // Office command was first opened.
    let font_size_pt = session
        .font_size_pt
        .filter(|value| {
            value.is_finite() && *value >= MIN_WORD_FONT_SIZE_PT && *value <= MAX_WORD_FONT_SIZE_PT
        })
        .or_else(|| {
            request.font_size_pt.filter(|value| {
                value.is_finite()
                    && *value >= MIN_WORD_FONT_SIZE_PT
                    && *value <= MAX_WORD_FONT_SIZE_PT
            })
        })
        .or_else(|| {
            session
                .original_metadata
                .as_ref()
                .and_then(|metadata| metadata.font_size_pt)
                .filter(|value| {
                    value.is_finite()
                        && *value >= MIN_WORD_FONT_SIZE_PT
                        && *value <= MAX_WORD_FONT_SIZE_PT
                })
        })
        .unwrap_or(WORD_REFERENCE_FONT_SIZE_PT);
    let baseline = export
        .baseline
        .filter(|value| value.is_finite() && *value >= 0.0 && *value <= export.height)
        .ok_or_else(|| {
            "Word formula export is missing a valid mathematical baseline".to_string()
        })?;
    // Never silently fall back to the image bottom for a missing baseline. That
    // makes simple formulas look acceptable while subscripts, fractions and
    // tall delimiters are inserted at visibly different vertical positions on
    // Word versions that round InlineShape.Font.Position differently. Every
    // current renderer provides a mathematical baseline, so absence means the
    // artifact is incomplete and must not be committed with corrupt geometry.
    // Document import and edit replacement share this exact 14 pt reference
    // geometry path, including maximum width, descent and baseline rounding.
    calculate_word_svg_geometry_for_font(
        export.width,
        export.height,
        baseline,
        font_size_pt,
        Some(session.formula_letter_font.as_str()),
    )
}

/// A picture formula VisualTeX for Windows drew (its metadata has renderFontSizePt; the Mac's never does) that has
/// none of the Mac's scale tags reaches the app at the add-in's default size: its own size is the one its render
/// size was drawn at, scaled by the picture's height.
pub(super) fn windows_picture_font_size(
    powerpoint: Option<&MacOfflinePowerPointRequest>,
    metadata: Option<&VisualTeXFormulaMetadata>,
) -> Option<f64> {
    let powerpoint = powerpoint?;
    let metadata = metadata?;
    let assumed = powerpoint.font_size_pt?;
    if (assumed - DEFAULT_POWERPOINT_FONT_SIZE_PT).abs() > 0.001 {
        return None;
    }
    let render_font_size = metadata.render_font_size_pt.filter(|value| value.is_finite() && *value > 0.0)?;
    let natural_height = metadata.render_height_px.filter(|value| value.is_finite() && *value > 0.0)? * 0.75;
    let size = (render_font_size * powerpoint.height / natural_height * 100.0).round() / 100.0;
    (size.is_finite() && (MIN_POWERPOINT_FONT_SIZE_PT..=MAX_POWERPOINT_FONT_SIZE_PT).contains(&size)).then_some(size)
}

pub(super) fn calculate_powerpoint_geometry(
    request: &MacOfflinePowerPointRequest,
    session: &OfficeFormulaSession,
) -> Result<PowerPointGeometry, String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "PowerPoint Session has no formula export".to_string())?;
    if !export.width.is_finite()
        || !export.height.is_finite()
        || export.width <= 0.0
        || export.height <= 0.0
    {
        return Err("PowerPoint formula export has invalid dimensions".to_string());
    }

    // MathJax exports at a stable 14 pt reference size. The imported SVG keeps
    // its vector paths, so a PowerPoint point size is represented by uniformly
    // scaling the natural SVG bounds rather than rasterizing or stretching it.
    let reference_width_pt = export.width * 0.75;
    let reference_height_pt = export.height * 0.75;
    if !reference_width_pt.is_finite()
        || !reference_height_pt.is_finite()
        || reference_width_pt <= 0.0
        || reference_height_pt <= 0.0
    {
        return Err("PowerPoint SVG reference geometry is invalid".to_string());
    }

    let original = session.original_metadata.as_ref();
    let previous_reference_height = request
        .reference_height_pt
        .filter(|value| value.is_finite() && *value > 0.0)
        .or_else(|| {
            original
                .and_then(|metadata| metadata.reference_height_pt)
                .filter(|value| value.is_finite() && *value > 0.0)
        })
        .or_else(|| {
            original
                .and_then(|metadata| metadata.render_height_px)
                .filter(|value| value.is_finite() && *value > 0.0)
                .map(|value| value * 0.75)
        });

    let committed_font_size = session.font_size_pt.filter(|value| {
        value.is_finite()
            && *value >= MIN_POWERPOINT_FONT_SIZE_PT
            && *value <= MAX_POWERPOINT_FONT_SIZE_PT
    });
    let declared_font_size = request
        .font_size_pt
        .filter(|value| {
            value.is_finite()
                && *value >= MIN_POWERPOINT_FONT_SIZE_PT
                && *value <= MAX_POWERPOINT_FONT_SIZE_PT
        })
        .or_else(|| {
            original
                .and_then(|metadata| metadata.font_size_pt)
                .filter(|value| {
                    value.is_finite()
                        && *value >= MIN_POWERPOINT_FONT_SIZE_PT
                        && *value <= MAX_POWERPOINT_FONT_SIZE_PT
                })
        });

    // The actual selected shape height wins when a user resized an existing SVG
    // manually. This converts that physical height back to an equivalent point
    // size and prevents the next edit from jumping to a stale stored value.
    let observed_font_size = previous_reference_height
        .map(|height| request.height / height * POWERPOINT_REFERENCE_FONT_SIZE_PT)
        .filter(|value| {
            value.is_finite()
                && *value >= MIN_POWERPOINT_FONT_SIZE_PT
                && *value <= MAX_POWERPOINT_FONT_SIZE_PT
        });
    // The VBA adapter already converts a manually resized existing shape into
    // the Session's initial fontSizePt. After the editor opens, however, the
    // Session value is the user's explicit choice and must not be overwritten
    // by re-observing the old shape bounds during commit.
    let font_size_pt = committed_font_size
        .or(observed_font_size)
        .or(declared_font_size)
        .unwrap_or(DEFAULT_POWERPOINT_FONT_SIZE_PT);
    let point_scale = font_size_pt / POWERPOINT_REFERENCE_FONT_SIZE_PT;
    let target_width = reference_width_pt * point_scale;
    let target_height = reference_height_pt * point_scale;
    let center_x = request.left + request.width / 2.0;
    let center_y = request.top + request.height / 2.0;
    let left = center_x - target_width / 2.0;
    let top = center_y - target_height / 2.0;
    for (value, label) in [
        (left, "target left"),
        (top, "target top"),
        (target_width, "target width"),
        (target_height, "target height"),
    ] {
        validate_finite_geometry(value, label)?;
    }
    if target_width <= 0.0
        || target_height <= 0.0
        || target_width > 10_000.0
        || target_height > 10_000.0
    {
        return Err("PowerPoint target formula dimensions are invalid".to_string());
    }
    Ok(PowerPointGeometry {
        left,
        top,
        width: target_width,
        height: target_height,
        font_size_pt,
        reference_width_pt,
        reference_height_pt,
    })
}
