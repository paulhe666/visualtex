import { Configuration } from "mathjax-full/js/input/tex/Configuration.js";
import type { MmlNode } from "mathjax-full/js/core/MmlTree/MmlNode.js";
import type ParseOptions from "mathjax-full/js/input/tex/ParseOptions.js";

// Record style at token creation, where the real TeX group environment is
// authoritative. A later tree walk cannot distinguish an inner reset (including
// legacy {\\rm ...}) from an operator dictionary's default normal attribute.
const explicitVariants = new WeakMap<MmlNode, string>();
const basicVariants = new Set(["normal", "italic", "bold", "bold-italic"]);
export const EXPLICIT_MATH_FONT_PACKAGE = "visualtex-explicit-math-font";
Configuration.create(EXPLICIT_MATH_FONT_PACKAGE, {
  config: [(_configuration, jax) => {
    const factory = jax.parseOptions.nodeFactory;
    // Compose with the installed token creator, especially boldsymbol's flags.
    const createToken = factory.get("token");
    factory.set("token", (factory, kind, def = {}, text = "") => {
      const node = createToken(factory, kind, def, text);
      const font = factory.configuration.parser.stack.env.font;
      const variant = font === "-tex-mathit" ? "italic" : String(font ?? "");
      if (basicVariants.has(variant) &&
          (kind === "mo" || (kind === "mi" && /^\p{Sm}+$/u.test(text)))) {
        explicitVariants.set(node, variant);
      }
      return node;
    });
  }, 100],
  // Run before boldsymbol's default-priority pass: it can then combine weight
  // with explicit slant. Run after parsing, since Unicode Other overwrites the
  // token variant with the operator dictionary's normal default.
  postprocessors: [[({ data }: { data: ParseOptions }) => {
    data.root.walkTree((raw) => {
      const node = raw as MmlNode;
      const variant = explicitVariants.get(node);
      if (variant) node.attributes.set("mathvariant", variant);
    });
  }, -10]],
});
