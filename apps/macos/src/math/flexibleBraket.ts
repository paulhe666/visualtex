/**
 * braket.sty writes one balanced argument with internal | separators
 * (\\braket{\\phi|\\psi}); physics writes two arguments. The compatibility
 * macros only know the two-argument form, so rewrite the one-argument form to
 * the same delimiters before macro expansion or export.
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
      output += `\\left\\langle ${first.content}\\middle\\vert ${second.content}\\right\\rangle`;
      cursor = second.end;
      continue;
    }

    // braket.sty-style one-argument syntax: separators live inside the single
    // balanced argument. Replace only top-level unescaped | / || so nested groups
    // and explicit \\| remain untouched.
    let depth = 0;
    let body = "";
    let sawSeparator = false;
    for (let index = 0; index < first.content.length; index += 1) {
      const character = first.content[index];
      if (character === "\\") {
        body += character;
        if (index + 1 < first.content.length) body += first.content[++index];
        continue;
      }
      if (character === "{") {
        depth += 1;
        body += character;
        continue;
      }
      if (character === "}") {
        depth = Math.max(0, depth - 1);
        body += character;
        continue;
      }
      if (character === "|" && depth === 0) {
        sawSeparator = true;
        if (first.content[index + 1] === "|") {
          body += "\\middle\\Vert ";
          index += 1;
        } else {
          body += "\\middle\\vert ";
        }
        continue;
      }
      body += character;
    }
    if (!sawSeparator) {
      output += source.slice(match.index, first.end);
      cursor = first.end;
      continue;
    }
    output += `\\left\\langle ${body}\\right\\rangle`;
    cursor = first.end;
  }
  return output;
}
