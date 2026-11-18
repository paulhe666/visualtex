// WPS (Windows 12.1) draws an SVG <text> whose font-family names a font in quotes ("Times New Roman", serif) in a
// sans-serif fallback; written without quotes (Times New Roman, serif) it finds the font. CSS allows a family name
// without quotes when it is a run of identifiers, so exported SVGs write such names bare. Any other name keeps its
// quotes, and so do names that would read as keywords (serif, inherit, ...) without them.

const KEYWORDS = new Set([
  "serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "math", "emoji", "fangsong",
  "ui-serif", "ui-sans-serif", "ui-monospace", "ui-rounded",
  "inherit", "initial", "unset", "revert", "revert-layer", "default",
]);

// one CSS identifier: a letter, "_" or a non-ASCII character first (a "-" may lead when another of those follows),
// then letters, digits, "-", "_" or non-ASCII characters
const IDENTIFIER = /^-?[A-Za-z_\u0080-\uFFFF][A-Za-z0-9_\u0080-\uFFFF-]*$/;

function canStandBare(name: string) {
  const words = name.split(" ");
  return words.every((word) => IDENTIFIER.test(word)) && !KEYWORDS.has(name.toLowerCase());
}

/** A CSS font-family list with the quotes taken off every name that may stand without them. */
export function bareFontFamilyList(list: string): string {
  const names: { text: string; quoted: boolean }[] = [];
  let current = "";
  let quote: string | null = null;
  let quoted = false;
  for (const character of list) {
    if (quote) {
      if (character === quote) quote = null;
      else current += character;
    } else if (character === '"' || character === "'") {
      quote = character;
      quoted = true;
    } else if (character === ",") {
      names.push({ text: current.trim(), quoted });
      current = "";
      quoted = false;
    } else {
      current += character;
    }
  }
  if (quote) return list; // unbalanced quotes: leave the list as it is
  names.push({ text: current.trim(), quoted });
  return names
    .filter((name) => name.text.length > 0)
    .map((name) => {
      const text = name.text.replace(/\s+/g, " ");
      if (!name.quoted || canStandBare(text)) return text;
      return `"${text.replace(/["\\]/g, (c) => `\\${c}`)}"`;
    })
    .join(", ");
}

/** The SVG with every font-family attribute's names written bare where CSS allows it. */
export function bareSvgFontFamilies(svg: string): string {
  return svg.replace(
    /\bfont-family=(?:"([^"]*)"|'([^']*)')/g,
    (whole, doubleQuoted: string | undefined, singleQuoted: string | undefined) => {
      const delimiter = doubleQuoted !== undefined ? '"' : "'";
      const decoded = (doubleQuoted ?? singleQuoted ?? "")
        .replace(/&quot;/g, '"')
        .replace(/&apos;|&#39;/g, "'")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&amp;/g, "&");
      const bare = bareFontFamilyList(decoded);
      if (bare === decoded) return whole;
      let escaped = bare.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
      escaped = delimiter === '"' ? escaped.replace(/"/g, "&quot;") : escaped.replace(/'/g, "&apos;");
      return `font-family=${delimiter}${escaped}${delimiter}`;
    },
  );
}
