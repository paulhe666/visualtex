export interface FormulaEquationTagSplit {
  latex: string;
  equationTag: string | null;
}

function matchingClosingBrace(source: string, openIndex: number) {
  let depth = 0;
  for (let index = openIndex; index < source.length; index += 1) {
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
  return -1;
}

/**
 * Extract a trailing equation tag from an editable formula body.
 *
 * Both canonical `\\tag{4.8.4}` and the malformed-but-common
 * `\\tag 4.8.4` form are accepted so older batch-import metadata can be
 * repaired without exposing the tag command to MathLive.
 */
export function splitFormulaEquationTag(source: string): FormulaEquationTagSplit {
  const trimmed = source.trim();
  if (!trimmed) return { latex: "", equationTag: null };

  for (let index = trimmed.lastIndexOf("\\tag"); index >= 0; index = trimmed.lastIndexOf("\\tag", index - 1)) {
    let cursor = index + "\\tag".length;
    if (trimmed[cursor] === "*") cursor += 1;
    const boundary = trimmed[cursor];
    if (boundary && boundary !== "{" && !/\s/.test(boundary)) continue;
    while (/\s/.test(trimmed[cursor] ?? "")) cursor += 1;
    if (cursor >= trimmed.length) continue;

    if (trimmed[cursor] === "{") {
      const close = matchingClosingBrace(trimmed, cursor);
      if (close < 0 || trimmed.slice(close + 1).trim()) continue;
      const equationTag = trimmed.slice(cursor + 1, close).trim();
      if (!equationTag) continue;
      return {
        latex: trimmed.slice(0, index).trimEnd(),
        equationTag,
      };
    }

    const equationTag = trimmed.slice(cursor).trim();
    if (!equationTag || /[{}]/.test(equationTag)) continue;
    return {
      latex: trimmed.slice(0, index).trimEnd(),
      equationTag,
    };
  }

  return { latex: trimmed, equationTag: null };
}

export function attachFormulaEquationTag(
  source: string,
  equationTag: string | null | undefined,
) {
  const body = splitFormulaEquationTag(source).latex;
  const tag = equationTag?.trim();
  return tag ? `${body}\\tag{${tag}}` : body;
}
