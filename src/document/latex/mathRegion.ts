// The formula around a place in LaTeX source ($…$, \(…\), \[…\], $$…$$, equation/align/gather/multline …), so
// it can be opened in the formula editor and its LaTeX written back.

export interface MathRegion {
  /** the whole formula, delimiters included */
  from: number;
  to: number;
  /** its LaTeX, without delimiters */
  innerFrom: number;
  innerTo: number;
  display: boolean;
  /** the environment (equation, align*, …) or "" for $, \(, \[ */
  env: string;
}

const ENVS = /^(equation|align|alignat|flalign|gather|multline|eqnarray|math|displaymath)\*?$/;

/** All formulas of a text, outside comments and verbatim environments. */
export function mathRegions(text: string): MathRegion[] {
  const out: MathRegion[] = [];
  let i = 0;
  const n = text.length;
  const skipComment = () => {
    while (i < n && text[i] !== "\n") i++;
  };
  while (i < n) {
    const c = text[i];
    if (c === "%") {
      skipComment();
      continue;
    }
    if (c === "\\") {
      const rest = text.slice(i, i + 40);
      const verb = /^\\begin\{(verbatim\*?|lstlisting|minted|comment|Verbatim)\}/.exec(rest);
      if (verb) {
        const end = text.indexOf(`\\end{${verb[1]}}`, i);
        i = end < 0 ? n : end + verb[1].length + 6;
        continue;
      }
      if (rest.startsWith("\\verb")) {
        const delim = text[i + 5];
        const end = text.indexOf(delim, i + 6);
        i = end < 0 ? n : end + 1;
        continue;
      }
      const env = /^\\begin\{([A-Za-z]+\*?)\}/.exec(rest);
      if (env && ENVS.test(env[1])) {
        const close = `\\end{${env[1]}}`;
        const end = text.indexOf(close, i);
        if (end >= 0) {
          out.push({ from: i, to: end + close.length, innerFrom: i + env[0].length, innerTo: end, display: true, env: env[1] });
          i = end + close.length;
          continue;
        }
      }
      if (rest.startsWith("\\[") || rest.startsWith("\\(")) {
        const close = rest[1] === "[" ? "\\]" : "\\)";
        const end = findClose(text, i + 2, close);
        if (end >= 0) {
          out.push({ from: i, to: end + 2, innerFrom: i + 2, innerTo: end, display: rest[1] === "[", env: "" });
          i = end + 2;
          continue;
        }
      }
      i += 2;
      continue;
    }
    if (c === "$") {
      const double = text[i + 1] === "$";
      const open = double ? 2 : 1;
      const end = findClose(text, i + open, double ? "$$" : "$");
      if (end >= 0) {
        out.push({ from: i, to: end + open, innerFrom: i + open, innerTo: end, display: double, env: "" });
        i = end + open;
        continue;
      }
    }
    i++;
  }
  return out;
}

/** The closing delimiter after `from` (escaped characters skipped; a blank line ends inline math). */
function findClose(text: string, from: number, close: string): number {
  for (let j = from; j < text.length; j++) {
    if (text[j] === "\\") {
      if (text.startsWith(close, j) && close.startsWith("\\")) return j;
      j++;
      continue;
    }
    if (close === "$" && text[j] === "\n" && text[j + 1] === "\n") return -1;
    if (text.startsWith(close, j)) return j;
  }
  return -1;
}

export function mathRegionAt(text: string, pos: number): MathRegion | null {
  for (const r of mathRegions(text)) {
    if (pos >= r.from && pos <= r.to) return r;
    if (r.from > pos) break;
  }
  return null;
}

const ALIGNED: Record<string, string> = { align: "aligned", flalign: "aligned", alignat: "aligned", eqnarray: "aligned", gather: "gathered", multline: "gathered" };

/**
 * The formula as the formula editor takes it (labels and \nonumber set aside; align/gather rows in
 * aligned/gathered), and how its edited LaTeX goes back into the source.
 */
export function formulaForEditor(text: string, region: MathRegion): { latex: string; back: (edited: string) => string } {
  let inner = text.slice(region.innerFrom, region.innerTo);
  const kept: string[] = [];
  inner = inner.replace(/\\(label\{[^{}]*\}|nonumber\b|notag\b|tag\*?\{[^{}]*\})/g, (whole) => {
    kept.push(whole);
    return "";
  });
  const base = region.env.replace(/\*$/, "");
  const wrap = ALIGNED[base];
  const trimmed = inner.trim();
  const latex = wrap && /\\\\|&/.test(trimmed) ? `\\begin{${wrap}}${trimmed}\\end{${wrap}}` : trimmed;
  const leading = /^\s*/.exec(inner)![0];
  const trailing = /\s*$/.exec(inner)![0];
  return {
    latex,
    back: (edited) => {
      let body = edited.trim();
      const unwrap = wrap && new RegExp(`^\\\\begin\\{${wrap}\\}([\\s\\S]*)\\\\end\\{${wrap}\\}$`).exec(body);
      if (unwrap) body = unwrap[1].trim();
      if (kept.length) body += " " + kept.join(" ");
      return (region.display && region.env ? leading || "\n" : leading) + body + (region.display && region.env ? trailing || "\n" : trailing);
    },
  };
}
