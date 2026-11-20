/**
 * LaTeX handed to other programs spells bold math as \boldsymbol.
 *
 * The editor keeps bold-italic as \mathbfit (and accepts \bm), which KaTeX
 * (Typora, Obsidian, most Markdown editors) cannot render and pdfLaTeX only
 * knows with extra packages. \boldsymbol renders in KaTeX, MathJax and
 * pdfLaTeX alike. Only text leaving the app is rewritten; formulas, the
 * source panel and the Office conversions keep the editor's own spelling.
 */

const boldCommands = new Set(["mathbfit", "bm", "boldsymbol", "mathbf"]);

// \boldsymbol keeps capital Greek upright; \mathbfit makes it italic.
const uprightGreekCapital =
  /\\(?:var)?(?:Gamma|Delta|Theta|Lambda|Xi|Pi|Sigma|Upsilon|Phi|Psi|Omega)(?![A-Za-z])/;

function groupEnd(source: string, open: number): number | null {
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    const character = source[index];
    if (character === "\\") {
      index += 1;
      continue;
    }
    if (character === "{") depth += 1;
    else if (character === "}") {
      depth -= 1;
      if (depth === 0) return index;
    }
  }
  return null;
}

/** The content of `\name{...}` when the text is exactly that one group. */
function singleGroup(text: string, name: string): string | null {
  const prefix = `\\${name}{`;
  if (!text.startsWith(prefix)) return null;
  const end = groupEnd(text, prefix.length - 1);
  return end === text.length - 1 ? text.slice(prefix.length, end) : null;
}

function rewriteBold(command: string, argument: string): string {
  // Nested bold commands: the innermost one decides, as in MathJax.
  const innerBoldSymbol = singleGroup(argument, "boldsymbol");
  if (innerBoldSymbol !== null) return argument;
  if (command === "mathbf") return `\\mathbf{${argument}}`;
  if (command === "mathbfit" && uprightGreekCapital.test(argument)) {
    return `\\boldsymbol{\\mathit{${argument}}}`;
  }
  return `\\boldsymbol{${argument}}`;
}

export function toPortableBoldLatex(latex: string): string {
  const source = String(latex ?? "");
  if (!/\\(?:mathbfit|bm)(?![A-Za-z])/.test(source)) return source;
  let result = "";
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (character !== "\\") {
      result += character;
      continue;
    }
    const name = /^[A-Za-z]+/.exec(source.slice(index + 1))?.[0];
    if (!name) {
      // An escaped character such as \{ or \\ is copied as one unit.
      result += source.slice(index, index + 2);
      index += 1;
      continue;
    }
    const afterName = index + 1 + name.length;
    if (!boldCommands.has(name)) {
      result += `\\${name}`;
      index = afterName - 1;
      continue;
    }
    let open = afterName;
    while (source[open] === " ") open += 1;
    const end = source[open] === "{" ? groupEnd(source, open) : null;
    if (end === null) {
      // An unbraced argument (\bm x) takes one token either way.
      result += name === "mathbfit" || name === "bm" ? "\\boldsymbol" : `\\${name}`;
      index = afterName - 1;
      continue;
    }
    result += rewriteBold(name, toPortableBoldLatex(source.slice(open + 1, end)));
    index = end;
  }
  return result;
}
