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
  idotsint: "\\int\\!\\cdots\\!\\int",
  dotsint: "\\int\\!\\cdots\\!\\int",
  sqint: "⨖",
  sqiint: "⨖⨖",
  ointclockwise: "∱",
  varointctrclockwise: "∳",
  varoiint: "∯",
  landupint: "⨛",
  landdownint: "⨜",
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

export const EXTENDED_INTEGRAL_MATHML_MACROS = {
  ...EXTENDED_INTEGRAL_SYMBOLS,
} as const;

export const EXTENDED_INTEGRAL_SVG_MACROS = {
  oiint: svgMarker("oiint", "\\iint"),
  oiiint: svgMarker("oiiint", "\\iiint"),
  ...RARE_INTEGRAL_SVG_MACROS,
  ...ESINT_INTEGRAL_SVG_MACROS,
} as const;
