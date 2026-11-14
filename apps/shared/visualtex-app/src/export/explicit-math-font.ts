import { Configuration } from "mathjax-full/js/input/tex/Configuration.js";
import { TEXCLASS, type MmlNode } from "mathjax-full/js/core/MmlTree/MmlNode.js";
import { MmlMi } from "mathjax-full/js/core/MmlTree/MmlNodes/mi.js";
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
      const environment = factory.configuration.parser.stack.env;
      const font = environment.font;
      const variant = font === "-tex-mathit" ? "italic" : String(font ?? "");
      const declared = String(def.mathvariant ?? "");
      // Explicit basic styles apply to a mathematical identifier regardless of
      // whether it was typed literally or obtained from a fixed-family command.
      // Preserve intrinsic alphabets and TeX text-italic, and keep operator-name
      // contents under their own parser authority.
      const identifier = kind === "mi" && MmlMi.singleCharacter.test(text) &&
        !environment.operatorLetters && (!declared || basicVariants.has(declared));
      if (basicVariants.has(variant) && (kind === "mo" || identifier)) {
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
      if (variant && !(node.kind === "mi" && node.getProperty("texClass") === TEXCLASS.OP))
        node.attributes.set("mathvariant", variant);
    });
  }, -10]],
});
