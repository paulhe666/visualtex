/**
 * braket.sty writes one balanced argument with internal | separators
 * (\\braket{\\phi|\\psi}); physics writes two arguments, the second one only
 * when a brace group follows, and one argument without a | of its own is
 * physics' too (\\braket{a} = ⟨a|a⟩), as the editor's kernel reads them. The
 * compatibility macros only know fixed arguments, so rewrite the forms to
 * their delimiters before macro expansion or export.
 */
type BalancedArgument = {
  content: string;
  end: number;
};

function skipLatexWhitespace(source: string, index: number) {
  while (index < source.length && /\s/.test(source[index])) index += 1;
  return index;
}
function readBalancedArgument(
  source: string,
  start: number,
  open: string,
  close: string,
): BalancedArgument | null {
  if (source[start] !== open) return null;
  if (open === close) {
    for (let index = start + 1; index < source.length; index += 1) {
      if (source[index] === "\\") {
        index += 1;
        continue;
      }
      if (source[index] === close) {
        return { content: source.slice(start + 1, index), end: index + 1 };
      }
    }
    return null;
  }

  let depth = 0;
  for (let index = start; index < source.length; index += 1) {
    if (source[index] === "\\") {
      index += 1;
      continue;
    }
    if (source[index] === open) depth += 1;
    else if (source[index] === close) {
      depth -= 1;
      if (depth === 0) {
        return { content: source.slice(start + 1, index), end: index + 1 };
      }
    }
  }
  return null;
}

function readPackageMacroArgument(source: string, start: number) {
  const position = skipLatexWhitespace(source, start);
  if (source[position] === "{") {
    return readBalancedArgument(source, position, "{", "}");
  }
  if (source[position] === "\\") {
    const command = source.slice(position).match(/^\\(?:[A-Za-z@]+|.)/);
    return command
      ? { content: command[0], end: position + command[0].length }
      : null;
  }
  return position < source.length
    ? { content: source[position], end: position + 1 }
    : null;
}

/** Whether braket.sty's argument has a | of its own (not inside braces). */
function hasTopLevelBar(text: string) {
  let depth = 0;
  for (let index = 0; index < text.length; index += 1) {
    if (text[index] === "\\") {
      index += 1;
      continue;
    }
    if (text[index] === "{") depth += 1;
    else if (text[index] === "}") depth -= 1;
    else if (text[index] === "|" && depth === 0) return true;
  }
  return false;
}

export function normalizeFlexiblePackageLatexCommands(
  source: string,
  excludedCommands: ReadonlySet<string> = new Set(),
) {
  if (excludedCommands.has("braket")) return source;
  const commandPattern = /\\braket(?![A-Za-z@])/g;
  let output = "";
  let cursor = 0;
  while (cursor < source.length) {
    commandPattern.lastIndex = cursor;
    const match = commandPattern.exec(source);
    if (!match) {
      output += source.slice(cursor);
      break;
    }
    output += source.slice(cursor, match.index);
    const first = readPackageMacroArgument(source, commandPattern.lastIndex);
    if (!first) {
      output += match[0];
      cursor = commandPattern.lastIndex;
      continue;
    }

    const secondStart = skipLatexWhitespace(source, first.end);
    const second = source[secondStart] === "{"
      ? readBalancedArgument(source, secondStart, "{", "}")
      : null;
    if (second) {
      const bra = normalizeFlexiblePackageLatexCommands(first.content);
      const ket = normalizeFlexiblePackageLatexCommands(second.content);
      output += `\\left\\langle ${bra}\\middle\\vert ${ket}\\right\\rangle${/^[A-Za-z]/.test(source[second.end] ?? "") ? " " : ""}`;
      cursor = second.end;
      continue;
    }

    // braket.sty's one argument with | separators of its own (physics'
    // second argument is only a brace group, never the next token): the
    // separators stay as written. Without one, physics' \\braket{a} = ⟨a|a⟩.
    const content = normalizeFlexiblePackageLatexCommands(first.content);
    output += hasTopLevelBar(first.content)
      ? `\\mathinner{\\langle{${content}}\\rangle}`
      : `\\left\\langle ${content}\\middle\\vert ${content}\\right\\rangle${/^[A-Za-z]/.test(source[first.end] ?? "") ? " " : ""}`;
    cursor = first.end;
  }
  return output;
}
