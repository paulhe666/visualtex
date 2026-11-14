import type {AbstractMmlTokenNode, MmlNode} from "mathjax-full/js/core/MmlTree/MmlNode.js";

/** The subset of FontData used by the MathML rendering projection. */
interface GlyphFontData {
  getVariant(name: string): {chars: Record<number, unknown>} | undefined;
  getChar(name: string, scalar: number): unknown;
}

const BASIC_VARIANTS = ["normal", "bold", "italic", "bold-italic"] as const;

/**
 * MathJax's math alphabets inherit missing operator glyphs from another variant.
 * Export that actual rendering decision, rather than asking native consumers
 * for a nonexistent calligraphic/fraktur/etc. symbol. This never guesses from
 * a character name or a formula: identity of the renderer's resolved glyph
 * record proves the fallback. Real alphabet glyphs and explicit basic weight/
 * slant requests are not changed. No extra SVG rendering or font I/O is needed.
 */
export function materializeEffectiveOperatorVariants(root: MmlNode, font: GlyphFontData): void {
  root.walkTree((raw) => {
    const node = raw as MmlNode;
    if (node.kind !== "mo") return;
    const requested = String(node.attributes.get("mathvariant") ?? "normal");
    if ((BASIC_VARIANTS as readonly string[]).includes(requested)) return;
    const text = (node as AbstractMmlTokenNode).getText();
    const scalars = Array.from(text);
    if (scalars.length !== 1 || /[\p{L}\p{N}]/u.test(text)) return;
    const scalar = text.codePointAt(0)!;
    const internal = String(node.attributes.get("data-mjx-variant") ?? requested);
    const variant = font.getVariant(internal);
    if (!variant || Object.hasOwn(variant.chars, scalar)) return;
    const glyph = font.getChar(internal, scalar);
    if (!glyph) return;
    const effective = BASIC_VARIANTS.find((name) => font.getChar(name, scalar) === glyph);
    if (effective) node.attributes.set("mathvariant", effective);
  });
}
