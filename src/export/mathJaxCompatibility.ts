import type {
  SVGCharData,
  SVGFontData,
} from "mathjax-full/js/output/svg/FontData.js";
import {
  OIINT_SIZE1_OVAL_PATH,
  OIINT_SIZE1_OVAL_WIDTH_EM,
  OIINT_SIZE2_OVAL_PATH,
  OIINT_SIZE2_OVAL_WIDTH_EM,
  OIIINT_SIZE1_OVAL_PATH,
  OIIINT_SIZE1_OVAL_WIDTH_EM,
  OIIINT_SIZE2_OVAL_PATH,
  OIIINT_SIZE2_OVAL_WIDTH_EM,
} from "../math/integralGlyphs.ts";
import {
  RARE_INTEGRAL_GLYPHS,
  RARE_INTEGRAL_GLYPH_UNITS_PER_EM,
  type RareIntegralGlyphVariant,
} from "../math/rareIntegralGlyphs.generated.ts";

const MATHJAX_UNSUPPORTED_NARY_COMMANDS = {
  oiint: "∯",
  oiiint: "∰",
  intclockwise: "∱",
  varointclockwise: "∲",
  ointctrclockwise: "∳",
  sumint: "⨋",
  intbar: "⨍",
  intBar: "⨎",
  fint: "⨏",
  cirfnint: "⨐",
  awint: "⨑",
  intctrclockwise: "⨑",
  rppolint: "⨒",
  scpolint: "⨓",
  npolint: "⨔",
  pointint: "⨕",
  quatint: "⨖",
  intlarhk: "⨗",
  intx: "⨘",
  intcap: "⨙",
  intcup: "⨚",
  upint: "⨛",
  lowint: "⨜",
} as const;

export const MATHJAX_INTEGRAL_OPERATOR_CHARACTERS = [
  "∫",
  "∬",
  "∭",
  "∮",
  "∯",
  "∰",
  "∱",
  "∲",
  "∳",
  "⨋",
  "⨌",
  "⨍",
  "⨎",
  "⨏",
  "⨐",
  "⨑",
  "⨒",
  "⨓",
  "⨔",
  "⨕",
  "⨖",
  "⨗",
  "⨘",
  "⨙",
  "⨚",
  "⨛",
  "⨜",
] as const;

export const MATHJAX_UNSUPPORTED_NARY_COMMAND_NAMES = Object.freeze(
  Object.keys(MATHJAX_UNSUPPORTED_NARY_COMMANDS),
);

const unsupportedNaryCommandPattern = new RegExp(
  String.raw`(?<!\\)\\(${[...MATHJAX_UNSUPPORTED_NARY_COMMAND_NAMES]
    .sort((left, right) => right.length - left.length)
    .join("|")})(?![A-Za-z])`,
  "g",
);

/**
 * MathLive accepts several Unicode-math integral commands that MathJax 3.2
 * does not register in its TeX input jax. Replace only those command tokens
 * with their canonical operator characters before rendering. The editor and
 * formula metadata continue to retain the original portable LaTeX command.
 */
export function normalizeMathJaxUnsupportedNaryCommands(source: string) {
  return source.replace(
    unsupportedNaryCommandPattern,
    (_match, command: keyof typeof MATHJAX_UNSUPPORTED_NARY_COMMANDS) =>
      MATHJAX_UNSUPPORTED_NARY_COMMANDS[command],
  );
}

type ContourIntegralGlyphDefinition = {
  variant: "-smallop" | "-largeop";
  targetCodePoint: number;
  baseCodePoint: number;
  ovalPath: string;
  visualWidthEm: number;
};

const CONTOUR_INTEGRAL_GLYPHS: ContourIntegralGlyphDefinition[] = [
  {
    variant: "-smallop",
    targetCodePoint: 0x222f,
    baseCodePoint: 0x222c,
    ovalPath: OIINT_SIZE1_OVAL_PATH,
    visualWidthEm: OIINT_SIZE1_OVAL_WIDTH_EM,
  },
  {
    variant: "-largeop",
    targetCodePoint: 0x222f,
    baseCodePoint: 0x222c,
    ovalPath: OIINT_SIZE2_OVAL_PATH,
    visualWidthEm: OIINT_SIZE2_OVAL_WIDTH_EM,
  },
  {
    variant: "-smallop",
    targetCodePoint: 0x2230,
    baseCodePoint: 0x222d,
    ovalPath: OIIINT_SIZE1_OVAL_PATH,
    visualWidthEm: OIIINT_SIZE1_OVAL_WIDTH_EM,
  },
  {
    variant: "-largeop",
    targetCodePoint: 0x2230,
    baseCodePoint: 0x222d,
    ovalPath: OIIINT_SIZE2_OVAL_PATH,
    visualWidthEm: OIIINT_SIZE2_OVAL_WIDTH_EM,
  },
];

function defineContourIntegralGlyph(
  font: SVGFontData,
  definition: ContourIntegralGlyphDefinition,
) {
  const { variant, targetCodePoint, baseCodePoint, ovalPath, visualWidthEm } =
    definition;
  const baseGlyph = font.getChar(variant, baseCodePoint) as
    | SVGCharData
    | undefined;
  if (!baseGlyph || baseGlyph.length !== 4 || !baseGlyph[3].p) {
    throw new Error(
      `MathJax SVG font is missing the ${variant} base glyph U+${baseCodePoint
        .toString(16)
        .toUpperCase()}.`,
    );
  }

  const [height, depth, advanceWidth, baseOptions] = baseGlyph;
  const italicCorrection = baseOptions.ic ?? 0;
  // MathJax advances an operator by `width + italic correction`. The oval
  // geometry includes that correction, so cloning both base values preserves
  // exactly the same following-token position as \iint/\iiint.
  if (Math.abs(advanceWidth + italicCorrection - visualWidthEm) > 1e-9) {
    throw new Error(
      `MathJax SVG font geometry for U+${baseCodePoint
        .toString(16)
        .toUpperCase()} no longer matches the contour-integral overlay.`,
    );
  }

  font.defineChars(variant, {
    [targetCodePoint]: [
      height,
      depth,
      advanceWidth,
      { ...baseOptions, p: `${baseOptions.p}${ovalPath}` },
    ],
  });
}

/**
 * MathJax 3 recognizes U+222F/U+2230 as large operators but its TeX SVG font
 * has no paths for them, so it otherwise emits small, system-font `<text>`.
 * Compose deterministic Size1/Size2 paths from the existing double/triple
 * integral glyph and the matching contour oval before creating the document.
 */
export function registerMathJaxContourIntegralGlyphs(font: SVGFontData) {
  for (const definition of CONTOUR_INTEGRAL_GLYPHS) {
    defineContourIntegralGlyph(font, definition);
  }
}

type RareIntegralVariant = {
  name: "-smallop" | "-largeop";
  glyph: RareIntegralGlyphVariant;
};

function mathJaxRareIntegralCharData(
  variant: RareIntegralGlyphVariant,
): SVGCharData {
  const unitsPerEm = RARE_INTEGRAL_GLYPH_UNITS_PER_EM;
  return [
    variant.height / unitsPerEm,
    variant.depth / unitsPerEm,
    variant.advanceWidth / unitsPerEm,
    {
      ic: variant.italicCorrection / unitsPerEm,
      p: variant.mathJaxPath,
    },
  ];
}

/**
 * Register the normalized STIX Two Math outlines for the Unicode integral
 * operators absent from MathJax's TeX SVG font. Both text and display
 * variants are provided so MathJax can keep its native large-operator,
 * scripts, and limits layout without falling back to a system-font `<text>`.
 */
export function registerMathJaxRareIntegralGlyphs(font: SVGFontData) {
  const smallCharacters: Record<number, SVGCharData> = {};
  const largeCharacters: Record<number, SVGCharData> = {};

  for (const definition of RARE_INTEGRAL_GLYPHS) {
    const variants: RareIntegralVariant[] = [
      { name: "-smallop", glyph: definition.small },
      { name: "-largeop", glyph: definition.large },
    ];
    for (const variant of variants) {
      const target =
        variant.name === "-smallop" ? smallCharacters : largeCharacters;
      target[definition.codePoint] = mathJaxRareIntegralCharData(variant.glyph);
    }
  }

  font.defineChars("-smallop", smallCharacters);
  font.defineChars("-largeop", largeCharacters);
}

/** Register every VisualTeX-specific integral outline before conversion. */
export function registerMathJaxIntegralGlyphs(font: SVGFontData) {
  registerMathJaxContourIntegralGlyphs(font);
  registerMathJaxRareIntegralGlyphs(font);
}
