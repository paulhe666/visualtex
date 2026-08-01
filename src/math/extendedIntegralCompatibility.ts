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

export const EXTENDED_INTEGRAL_SYMBOLS = {
  oiint: "∯",
  oiiint: "∰",
  ...RARE_INTEGRAL_SYMBOLS,
} as const;

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
    svgMarker(command, command === "iiiint" ? "\\iiint" : "\\int"),
  ]),
) as Record<string, string>;

export const EXTENDED_INTEGRAL_MATHML_MACROS = {
  ...EXTENDED_INTEGRAL_SYMBOLS,
} as const;

export const EXTENDED_INTEGRAL_SVG_MACROS = {
  oiint: svgMarker("oiint", "\\iint"),
  oiiint: svgMarker("oiiint", "\\iiint"),
  ...RARE_INTEGRAL_SVG_MACROS,
} as const;
