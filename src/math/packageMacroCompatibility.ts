export interface VisualTexMathLiveMacroDefinition {
  def: string;
  args: number;
  expand: false;
  captureSelection: false;
}

export type VisualTexPackageMathJaxMacro = readonly [
  replacement: string,
  argumentCount: number,
];

const macro = (
  def: string,
  args: number,
): VisualTexMathLiveMacroDefinition => ({
  def,
  args,
  expand: false,
  captureSelection: false,
});

/**
 * Curated package macros whose common braced forms can be represented exactly
 * with standard LaTeX primitives. The command name is preserved in MathLive's
 * `latex` value while the expansion remains fully editable.
 *
 * Delimiter-star and optional-order variants from the original packages are
 * intentionally not emulated here: VisualTeX only advertises the forms below.
 */
export const VISUALTEX_MATHLIVE_PACKAGE_MACROS = Object.freeze({
  // bm
  bm: macro("\\boldsymbol{#1}", 1),

  // physics / physics-patch basic braced forms
  abs: macro("\\left\\lvert #1\\right\\rvert", 1),
  norm: macro("\\left\\lVert #1\\right\\rVert", 1),
  comm: macro("\\left[#1,#2\\right]", 2),
  acomm: macro("\\left\\{#1,#2\\right\\}", 2),
  pb: macro("\\left\\{#1,#2\\right\\}", 2),
  dv: macro("\\frac{\\mathrm{d}#1}{\\mathrm{d}#2}", 2),
  pdv: macro("\\frac{\\partial #1}{\\partial #2}", 2),
  dd: macro("\\mathrm{d}#1", 1),
  bra: macro("\\left\\langle #1\\right\\rvert", 1),
  ket: macro("\\left\\lvert #1\\right\\rangle", 1),
  braket: macro(
    "\\left\\langle #1\\middle\\vert #2\\right\\rangle",
    2,
  ),
  expval: macro("\\left\\langle #1\\right\\rangle", 1),
  mel: macro(
    "\\left\\langle #1\\middle\\vert #2\\middle\\vert #3\\right\\rangle",
    3,
  ),
  ketbra: macro(
    "\\left\\lvert #1\\middle\\rangle\\!\\middle\\langle #2\\right\\rvert",
    2,
  ),
  vb: macro("\\mathbf{#1}", 1),
  va: macro("\\vec{\\mathbf{#1}}", 1),
  vu: macro("\\mathbf{\\hat{#1}}", 1),
} satisfies Record<string, VisualTexMathLiveMacroDefinition>);

export const VISUALTEX_MATHJAX_PACKAGE_MACROS = Object.freeze(
  Object.fromEntries(
    Object.entries(VISUALTEX_MATHLIVE_PACKAGE_MACROS).map(
      ([name, definition]) =>
        [name, [definition.def, definition.args] as const] as const,
    ),
  ) as Record<string, VisualTexPackageMathJaxMacro>,
);

export const VISUALTEX_PACKAGE_MACRO_NAMES = Object.freeze(
  Object.keys(VISUALTEX_MATHLIVE_PACKAGE_MACROS),
);

export const VISUALTEX_PACKAGE_MACRO_PATTERN = new RegExp(
  `\\\\(?:${[...VISUALTEX_PACKAGE_MACRO_NAMES]
    .sort((left, right) => right.length - left.length)
    .join("|")})(?:\\s|\\{|$)`,
);
