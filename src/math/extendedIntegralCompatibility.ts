export const RARE_INTEGRAL_SYMBOLS = {
  intclockwise: "∱",
  varointclockwise: "∲",
  ointctrclockwise: "∳",
  sumint: "⨋",
  iiiint: "⨌",
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

export const ESINT_INTEGRAL_REPLACEMENTS = {
  // Commands supplied by the esint package but absent from MathLive's native
  // command table. Multi-glyph fallbacks are used only for semantic MathML /
  // Word conversion; editor and SVG rendering use the official esint10 paths.
  idotsint: "\\int\\!\\cdots\\!\\int",
  dotsint: "\\int\\!\\cdots\\!\\int",
  sqint: "⨖",
  sqiint: "⨖⨖",
  ointclockwise: "∱",
  varointctrclockwise: "∳",
  varoiint: "∯",
  landupint: "⨛",
  landdownint: "⨜",
  // Existing spellings whose editor glyph must also follow esint10.
  iiiint: "⨌",
  intclockwise: "∱",
  ointctrclockwise: "∳",
  varointclockwise: "∲",
  fint: "⨏",
} as const;

export const EXTENDED_INTEGRAL_SYMBOLS = {
  oiint: "∯",
  oiiint: "∰",
  ...RARE_INTEGRAL_SYMBOLS,
  ...ESINT_INTEGRAL_REPLACEMENTS,
} as const;

export const EXTENDED_INTEGRAL_COMMANDS = Object.freeze(
  Object.keys(EXTENDED_INTEGRAL_SYMBOLS),
);

export const EXTENDED_INTEGRAL_COMMAND_PATTERN_SOURCE = [
  ...EXTENDED_INTEGRAL_COMMANDS,
]
  .sort((left, right) => right.length - left.length)
  .join("|");

const CANONICAL_COMMAND_BY_CHARACTER = new Map<string, string>([
  ["∯", "oiint"],
  ["∰", "oiiint"],
  ["∱", "intclockwise"],
  ["∲", "varointclockwise"],
  ["∳", "ointctrclockwise"],
  ["⨋", "sumint"],
  ["⨌", "iiiint"],
  ["⨍", "intbar"],
  ["⨎", "intBar"],
  ["⨏", "fint"],
  ["⨐", "cirfnint"],
  // Keep the long-standing VisualTeX spelling as the serialization form;
  // MathLive also accepts the esint alias \\awint.
  ["⨑", "intctrclockwise"],
  ["⨒", "rppolint"],
  ["⨓", "scpolint"],
  ["⨔", "npolint"],
  ["⨕", "pointint"],
  ["⨖", "quatint"],
  ["⨗", "intlarhk"],
  ["⨘", "intx"],
  ["⨙", "intcap"],
  ["⨚", "intcup"],
  ["⨛", "upint"],
  ["⨜", "lowint"],
]);

const EXTENDED_INTEGRAL_CHARACTER_PATTERN = /[∯∰∱∲∳⨋-⨜]/gu;

/**
 * MathLive and old OLE objects can serialize an integral as its Unicode glyph.
 * Convert it back to one canonical LaTeX command before it reaches the editor,
 * metadata store, MathJax, or Word. A trailing command terminator prevents the
 * restored control word from consuming a following Latin variable.
 */
export function normalizeExtendedIntegralLatexCommands(source: string) {
  return source.replace(EXTENDED_INTEGRAL_CHARACTER_PATTERN, (character) => {
    const command = CANONICAL_COMMAND_BY_CHARACTER.get(character);
    return command ? `\\${command} ` : character;
  });
}

function svgMarker(command: string, placeholder: string) {
  return `\\class{visualtex-integral-export-${command}}{${placeholder}}`;
}

const RARE_INTEGRAL_SVG_MACROS = Object.fromEntries(
  Object.keys(RARE_INTEGRAL_SYMBOLS).map((command) => [
    command,
    // Do not reference \\iiiint from its own compatibility macro. A triple-
    // integral placeholder provides the closest built-in MathJax width for the
    // four-integral vector without recursive macro expansion.
    svgMarker(command, command === "iiiint" ? "\\iiint" : "\\int"),
  ]),
) as Record<string, string>;

const ESINT_SVG_PLACEHOLDERS: Record<
  keyof typeof ESINT_INTEGRAL_REPLACEMENTS,
  string
> = {
  idotsint: "\\iiint",
  dotsint: "\\iiint",
  sqint: "\\int",
  sqiint: "\\iint",
  ointclockwise: "\\oint",
  varointctrclockwise: "\\oint",
  varoiint: "\\iint",
  landupint: "\\int",
  landdownint: "\\int",
  iiiint: "\\iiint",
  intclockwise: "\\oint",
  ointctrclockwise: "\\oint",
  varointclockwise: "\\oint",
  fint: "\\int",
};

const ESINT_INTEGRAL_SVG_MACROS = Object.fromEntries(
  Object.entries(ESINT_SVG_PLACEHOLDERS).map(([command, placeholder]) => [
    command,
    svgMarker(command, placeholder),
  ]),
) as Record<keyof typeof ESINT_INTEGRAL_REPLACEMENTS, string>;

/** Semantic operators used by MathML and native Word OMML conversion. */
export const EXTENDED_INTEGRAL_MATHML_MACROS = {
  ...EXTENDED_INTEGRAL_SYMBOLS,
} as const;

/**
 * SVG/OLE rendering uses standard large-operator placeholders so MathJax lays
 * out limits and surrounding content with integral metrics. The export runtime
 * replaces the marked glyph with the same contour/STIX vector used by MathLive.
 */
export const EXTENDED_INTEGRAL_SVG_MACROS = {
  oiint: svgMarker("oiint", "\\iint"),
  oiiint: svgMarker("oiiint", "\\iiint"),
  ...RARE_INTEGRAL_SVG_MACROS,
  ...ESINT_INTEGRAL_SVG_MACROS,
} as const;
