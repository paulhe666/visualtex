// tensor.sty's \tensor[pre]{T}{post} and \indices{^a_b} for the MathJax export
// paths (MathML, SVG), which do not know them: each index on an empty base of
// its own, so the indices keep tensor.sty's staggering ({}^{a}{}_{b}). The
// saved LaTeX is left untouched; the editor and document pictures set them
// with the layout engine (shared/tex-layout latex-adapter.ts).

/** The end (exclusive) of the brace or bracket group opened at `start`, or -1. */
function groupEnd(source: string, start: number, open: string, close: string): number {
  let depth = 0;
  for (let i = start; i < source.length; i++) {
    const ch = source[i];
    if (ch === "\\") {
      i++;
      continue;
    }
    if (ch === open) depth++;
    else if (ch === close && --depth === 0) return i + 1;
  }
  return -1;
}

function skipSpaces(source: string, i: number): number {
  while (i < source.length && /\s/.test(source[i])) i++;
  return i;
}

/** An index string ("^a_{bc}^\mu") as staggered scripts on empty bases. */
function staggered(indices: string): string {
  let out = "";
  let i = 0;
  while (i < indices.length) {
    i = skipSpaces(indices, i);
    const mark = indices[i];
    if (mark !== "^" && mark !== "_") {
      i++;
      continue;
    }
    i = skipSpaces(indices, i + 1);
    let arg: string;
    if (indices[i] === "{") {
      const end = groupEnd(indices, i, "{", "}");
      if (end < 0) break;
      arg = indices.slice(i + 1, end - 1);
      i = end;
    } else {
      const m = /^\\(?:[A-Za-z@]+|.)|^./.exec(indices.slice(i));
      if (!m) break;
      arg = m[0];
      i += m[0].length;
    }
    out += `{}${mark}{${arg}}`;
  }
  return out;
}

export function normalizeTensorIndexCommands(source: string): string {
  const pattern = /\\(tensor|indices)\*?(?![A-Za-z@])/g;
  let out = "";
  let cursor = 0;
  for (let m = pattern.exec(source); m; m = pattern.exec(source)) {
    let i = skipSpaces(source, pattern.lastIndex);
    let pre = "";
    if (m[1] === "tensor" && source[i] === "[") {
      const end = groupEnd(source, i, "[", "]");
      if (end < 0) continue;
      pre = source.slice(i + 1, end - 1);
      i = skipSpaces(source, end);
    }
    let nucleus = "";
    if (m[1] === "tensor") {
      if (source[i] !== "{") continue;
      const end = groupEnd(source, i, "{", "}");
      if (end < 0) continue;
      nucleus = source.slice(i + 1, end - 1);
      i = skipSpaces(source, end);
    }
    if (source[i] !== "{") continue;
    const end = groupEnd(source, i, "{", "}");
    if (end < 0) continue;
    const post = normalizeTensorIndexCommands(source.slice(i + 1, end - 1));
    out += source.slice(cursor, m.index);
    out += m[1] === "tensor" ? `{${staggered(normalizeTensorIndexCommands(pre))}{${normalizeTensorIndexCommands(nucleus)}}${staggered(post)}}` : staggered(post);
    cursor = end;
    pattern.lastIndex = end;
  }
  return out + source.slice(cursor);
}
