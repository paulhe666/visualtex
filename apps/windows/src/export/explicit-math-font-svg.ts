import { SVGWrapperFactory } from "mathjax-full/js/output/svg/WrapperFactory.js";
import type { SVGWrapper, SVGWrapperClass } from "mathjax-full/js/output/svg/Wrapper.js";
import type { BBox } from "mathjax-full/js/util/BBox.js";

// CSS's conventional synthetic oblique angle. Apply only when the font's
// requested slanted glyph is provably the same record as its upright glyph.
const SLANT = Math.tan(14 * Math.PI / 180);
type WrapperConstructor = new (...args: any[]) => SVGWrapper<any, any, any>;

export function createExplicitMathFontSvgFactory(): SVGWrapperFactory<any, any, any> {
  const factory = new SVGWrapperFactory();
  for (const kind of ["mo", "mi"]) {
    const Base = factory.getNodeClass(kind) as unknown as WrapperConstructor;
    class ExplicitObliqueSymbol extends Base {
      private needsOblique(): boolean {
        const variant = String(this.node.attributes.get("mathvariant"));
        const upright = variant === "italic" ? "normal" : variant === "bold-italic" ? "bold" : null;
        const text = this.getText();
        if (!upright || !/^\p{Sm}+$/u.test(text)) return false;
        // MathJax may merge adjacent relation tokens; preserve their scope too.
        return Array.from(text).every((character) => {
          const scalar = character.codePointAt(0)!;
          const glyph = this.font.getChar(variant, scalar);
          return !!glyph && glyph === this.font.getChar(upright, scalar);
        });
      }
      protected computeBBox(bbox: BBox, recompute = false): void {
        super.computeBBox(bbox, recompute);
        if (this.needsOblique()) {
          // Reserve the shear's full horizontal extent, including descenders.
          // Baseline and vertical dimensions remain unchanged.
          bbox.w += SLANT * (Math.max(0, bbox.h) + Math.max(0, bbox.d));
        }
      }
      toSVG(parent: any): void {
        super.toSVG(parent);
        if (!this.needsOblique()) return;
        const children = [...this.adaptor.childNodes(this.element)];
        for (const child of children) this.adaptor.remove(child);
        const shift = SLANT * Math.max(0, this.getBBox().d);
        const group = this.svg("g", {
          "data-visualtex-synthetic-oblique": "true",
          transform: `matrix(1 0 ${SLANT} 1 ${this.fixed(shift)} 0)`,
        }, children);
        this.adaptor.append(this.element, group);
      }
    }
    factory.setNodeClass(kind, ExplicitObliqueSymbol as unknown as SVGWrapperClass);
  }
  return factory;
}
