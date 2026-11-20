import { VISUALTEX_CORE_MATHLIVE_MACROS } from "./coreLatexAliases";
import { VISUALTEX_PHYSICS_KERNEL_MACROS } from "./physicsKernelMacros";

export type VisualTexPackageCompatibilityMacro = readonly [
  replacement: string,
  argumentCount: number,
];

const coreAndPhysicsCompatibilityMacros = Object.fromEntries(
  [...VISUALTEX_CORE_MATHLIVE_MACROS, ...VISUALTEX_PHYSICS_KERNEL_MACROS].map(
    (definition) => [
      definition.name,
      [definition.def, definition.args] as const,
    ],
  ),
) as Record<string, VisualTexPackageCompatibilityMacro>;

/**
 * Curated package-style aliases whose supported braced forms can be rewritten
 * into standard LaTeX primitives before they reach the VisualTeX core.
 *
 * This table is renderer-independent compatibility data. Legacy development
 * oracles may adapt it to their own macro shape, but production code must not
 * depend on a third-party editor/renderer contract.
 */
export const VISUALTEX_PACKAGE_COMPATIBILITY_MACROS = Object.freeze({
  ...coreAndPhysicsCompatibilityMacros,

  // bm
  bm: ["\\boldsymbol{#1}", 1] as const,

  // physics / physics-patch basic braced forms
  abs: ["\\left\\lvert #1\\right\\rvert", 1] as const,
  norm: ["\\left\\lVert #1\\right\\rVert", 1] as const,
  comm: ["\\left[#1,#2\\right]", 2] as const,
  acomm: ["\\left\\{#1,#2\\right\\}", 2] as const,
  pb: ["\\left\\{#1,#2\\right\\}", 2] as const,
  dv: ["\\frac{\\mathrm{d}#1}{\\mathrm{d}#2}", 2] as const,
  pdv: ["\\frac{\\partial #1}{\\partial #2}", 2] as const,
  dd: ["\\mathrm{d}#1", 1] as const,
  bra: ["\\left\\langle #1\\right\\rvert", 1] as const,
  ket: ["\\left\\lvert #1\\right\\rangle", 1] as const,
  braket: ["\\left\\langle #1\\middle\\vert #2\\right\\rangle", 2] as const,
  expval: ["\\left\\langle #1\\right\\rangle", 1] as const,
  mel: ["\\left\\langle #1\\middle\\vert #2\\middle\\vert #3\\right\\rangle", 3] as const,
  ketbra: ["\\left\\lvert #1\\right\\rangle\\!\\left\\langle #2\\right\\rvert", 2] as const,
  vb: ["\\mathbf{#1}", 1] as const,
  va: ["\\vec{\\mathbf{#1}}", 1] as const,
  vu: ["\\mathbf{\\hat{#1}}", 1] as const,
} satisfies Record<string, VisualTexPackageCompatibilityMacro>);

export const VISUALTEX_PACKAGE_MACRO_NAMES = Object.freeze(
  Object.keys(VISUALTEX_PACKAGE_COMPATIBILITY_MACROS),
);

export const VISUALTEX_PACKAGE_MACRO_PATTERN = new RegExp(
  `\\\\(?:${[...VISUALTEX_PACKAGE_MACRO_NAMES]
    .sort((left, right) => right.length - left.length)
    .join("|")})(?:\\s|\\{|$)`,
);
