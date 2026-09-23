//! Non-barred two-argument forms. Logical ownership stays in Fraction; glyph
//! selection, vertical constraints and natural dimensions come from font data.
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::model::Node;
use crate::scene::Rect;
use crate::session::NodeId;
use super::{MathStyle, boxes::{Builder, Fragment}};

fn valid_distances(values: &[f64]) -> EngineResult<()> {
    if values.iter().any(|v| !v.is_finite() || *v < 0.0) {
        return Err(EngineError::new(ErrorCode::InvalidInput,
            "Font has invalid stack or skewed fraction distances"));
    }
    Ok(())
}
fn bounds(f: &Fragment) -> Rect {
    let logical = Rect::new(f.advance.min(0.0), -f.ascent, f.advance.max(0.0), f.descent);
    f.ink.map_or(logical, |ink| logical.union(ink))
}

impl Builder<'_> {
    pub(super) fn binomial(&mut self, id: NodeId, numerator: &Node,
        denominator: &Node, style: MathStyle, cramped: bool, depth: usize)
        -> EngineResult<Fragment>
    {
        let scale = style.scale(self.math);
        let num = self.node(numerator, style.fraction_child(), cramped, depth + 1)?;
        let den = self.node(denominator, style.fraction_child(), true, depth + 1)?;
        let m = self.math;
        let display = style == MathStyle::Display;
        let mut u = (if display { m.stack_top_display_shift } else { m.stack_top_shift }) * scale;
        let mut v = (if display { m.stack_bottom_display_shift } else { m.stack_bottom_shift }) * scale;
        let gap = (if display { m.stack_display_gap } else { m.stack_gap }) * scale;
        valid_distances(&[u, v, gap])?;
        // A zero-bar fraction uses Stack constants. The missing clearance is
        // distributed equally, rather than borrowing fraction-bar gaps.
        let extra = (gap - (u + v - num.descent - den.ascent)).max(0.0) / 2.0;
        u += extra; v += extra;
        let width = num.advance.abs().max(den.advance.abs());
        let nx = (width - num.advance) / 2.0;
        let dx = (width - den.advance) / 2.0;
        let mut f = Fragment::new(id, "fraction");
        f.child(num, nx, -u); f.child(den, dx, v); f.advance = width;

        // Both delimiters share the axis/target, but each uses its own designed
        // glyph or genuine assembly. Their glyphs belong to this Fraction node,
        // not duplicate same-ID nested layout regions.
        let axis = m.axis * scale;
        let target = 2.0 * (f.ascent - axis).max(f.descent + axis).max(1e-6);
        let (left_font, left) = self.vertical_product_or_math_glyph('(', crate::model::AtomClass::Open, target, scale)?;
        let (right_font, right) = self.vertical_product_or_math_glyph(')', crate::model::AtomClass::Close, target, scale)?;
        let inner = bounds(&f);
        // Keep ink overhang and negative glue inside, without changing child
        // source or assuming every subexpression has a positive advance.
        let shift = left.width * scale - inner.x_min;
        f.translate_x(shift);
        let right_x = shift + inner.x_max;
        for (font, shape, x) in [(left_font, left, 0.0), (right_font, right, right_x)] {
            let top = -axis - shape.height * scale / 2.0;
            f.advance = f.advance.max(x + shape.width * scale);
            for piece in shape.pieces { f.glyph(font, piece.shape, x, top + piece.y * scale, scale); }
        }
        f.italic = 0.0;
        Ok(f)
    }

    pub(super) fn bevelled_fraction(&mut self, id: NodeId, numerator: &Node,
        denominator: &Node, style: MathStyle, cramped: bool, depth: usize)
        -> EngineResult<Fragment>
    {
        let scale = style.scale(self.math);
        // nicefrac keeps its two editable arguments in reduced script style.
        // This finite product policy is not a claim to run arbitrary units.sty.
        let num = self.node(numerator, style.script_child(), cramped, depth + 1)?;
        let den = self.node(denominator, style.script_child(), true, depth + 1)?;
        let axis = self.math.axis * scale;
        let hgap = self.math.skewed_horizontal_gap * scale;
        let vgap = self.math.skewed_vertical_gap * scale;
        valid_distances(&[hgap, vgap])?;
        let nb = bounds(&num); let db = bounds(&den);
        let u = num.descent + axis + vgap / 2.0;
        let v = den.ascent - axis + vgap / 2.0;
        let above = num.ascent + u;
        let below = den.descent + v;
        let target = 2.0 * (above - axis).max(below + axis).max(1e-6);
        let (font, slash) = self.vertical_product_or_math_glyph('/', crate::model::AtomClass::Inner, target, scale)?;
        let slash_x = nb.width() + hgap / 2.0;
        let den_x = slash_x + slash.width * scale + hgap / 2.0 - db.x_min;
        let advance = den_x + db.x_max;
        let top = -axis - slash.height * scale / 2.0;
        let mut f = Fragment::new(id, "fraction");
        f.child(num, -nb.x_min, -u);
        f.child(den, den_x, v);
        for piece in slash.pieces { f.glyph(font, piece.shape, slash_x, top + piece.y * scale, scale); }
        f.advance = advance;
        f.italic = 0.0;
        Ok(f)
    }
}
