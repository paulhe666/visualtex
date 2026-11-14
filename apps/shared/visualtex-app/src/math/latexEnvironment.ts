const ENVIRONMENT_TOKEN_PATTERN = /\\(begin|end)\s*\{([^{}\r\n]+)\}/gu;

function isEscaped(source: string, index: number): boolean {
  let slashCount = 0;
  for (let cursor = index - 1; cursor >= 0 && source[cursor] === "\\"; cursor -= 1) {
    slashCount += 1;
  }
  return slashCount % 2 === 1;
}

function maskLatexComments(source: string): string {
  let masked = "";
  let inComment = false;

  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (inComment) {
      if (character === "\n" || character === "\r") {
        inComment = false;
        masked += character;
      } else {
        masked += " ";
      }
      continue;
    }

    if (character === "%" && !isEscaped(source, index)) {
      inComment = true;
      masked += " ";
      continue;
    }
    masked += character;
  }

  return masked;
}

/**
 * Removes one display-math delimiter pair only when it wraps the complete
 * source. Keeping this separate from environment parsing prevents callers from
 * accidentally handing MathJax literal `\\[` / `\\]` or `$$` tokens after an
 * editor document has already serialized an aligned environment for display.
 */
export function unwrapSingleLatexDisplayMath(source: string): string | null {
  const candidate = source.replace(/\r\n?/g, "\n").trim();
  if (candidate.startsWith("\\[") && candidate.endsWith("\\]")) {
    const inner = candidate.slice(2, -2).trim();
    return inner || null;
  }
  if (!candidate.startsWith("$$") || !candidate.endsWith("$$")) return null;

  const inner = candidate.slice(2, -2).trim();
  if (!inner) return null;
  for (let index = 0; index < inner.length - 1; index += 1) {
    if (
      inner[index] === "$" &&
      inner[index + 1] === "$" &&
      !isEscaped(inner, index)
    ) {
      return null;
    }
  }
  return inner;
}

/**
 * Returns true only when the source consists of one complete LaTeX environment.
 * Nested environments and comments are supported; trailing mathematical content
 * is deliberately rejected so ordinary multi-formula raw input can still split
 * into independent VisualTeX rows.
 */
export function isSingleCompleteLatexEnvironment(source: string): boolean {
  const candidate = maskLatexComments(source).trim();
  if (!candidate.startsWith("\\begin")) return false;

  const stack: string[] = [];
  let rootEnd = -1;
  ENVIRONMENT_TOKEN_PATTERN.lastIndex = 0;

  for (let match = ENVIRONMENT_TOKEN_PATTERN.exec(candidate); match; match = ENVIRONMENT_TOKEN_PATTERN.exec(candidate)) {
    if (isEscaped(candidate, match.index)) continue;
    const [, tokenKind, rawEnvironmentName] = match;
    const environmentName = rawEnvironmentName.trim();
    if (!environmentName) return false;

    if (stack.length === 0 && match.index !== 0) return false;
    if (tokenKind === "begin") {
      stack.push(environmentName);
      continue;
    }

    if (stack.length === 0 || stack.at(-1) !== environmentName) return false;
    stack.pop();
    if (stack.length === 0) {
      rootEnd = ENVIRONMENT_TOKEN_PATTERN.lastIndex;
      break;
    }
  }

  if (rootEnd < 0 || stack.length !== 0) return false;
  return candidate.slice(rootEnd).trim().length === 0;
}
