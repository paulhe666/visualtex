//! Root layout follows font MATH metrics and the MathML Core radical box rules.
//! https://www.w3.org/TR/mathml-core/#radicals-msqrt-mroot
use crate::error::EngineResult;
use crate::model::{AtomClass,Node};
use crate::scene::Rect;
use crate::session::NodeId;
use super::{MathStyle, boxes::{Builder, Fragment}};

impl Builder<'_> {
    pub fn radical(&mut self, id: NodeId, degree: Option<&Node>, radicand: &Node,
        style: MathStyle, depth: usize) -> EngineResult<Fragment> {
        let scale = style.scale(self.math);
        // The radicand remains at the ambient math size but uses cramped scripts.
        let base = self.node(radicand, style, true, depth + 1)?;
        let gap = if style == MathStyle::Display { self.math.radical_display_gap } else { self.math.radical_gap } * scale;
        let rule = self.math.radical_rule * scale;
        let extra = self.math.radical_extra_ascender * scale;
        let ink_ascent = base.ink.map_or(base.ascent, |r| -r.y_min);
        let ink_height = base.ink.map_or(base.ascent + base.descent, |r| r.height());
        let (surd_font,surd) = self.vertical_product_or_math_glyph(
            '√', AtomClass::Ordinary, ink_height + gap + rule, scale,
        )?;
        let ascent = base.ascent.max(ink_ascent + gap + rule + extra);
        let top = -ascent + extra;
        let descent = base.descent.max(top + surd.height * scale);
        let degree = degree.map(|d| self.node(d, MathStyle::ScriptScript, true, depth + 1)).transpose()?;
        let (before, surd_x) = degree.as_ref().map_or((0.0, 0.0), |d| {
            let before = (self.math.radical_kern_before_degree * scale).max(0.0);
            let after = (self.math.radical_kern_after_degree * scale).max(-d.advance);
            (before, before + d.advance + after)
        });
        let base_x = surd_x + surd.width * scale;
        let mut result = Fragment::new(id, "radical");
        result.ascent = ascent;
        result.descent = descent;
        result.advance = base_x + base.advance;
        for piece in surd.pieces {
            result.glyph(surd_font, piece.shape, surd_x, top + piece.y * scale, scale);
        }
        result.rule(Rect::new(base_x, top, base_x + base.advance, top + rule));
        result.child(base, base_x, 0.0);
        if let Some(degree) = degree {
            let y = descent - self.math.radical_degree_raise * (ascent + descent) - degree.descent;
            result.child(degree, before, y);
        }
        Ok(result)
    }
}
