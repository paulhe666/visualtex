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

export function isSingleCompleteLatexEnvironment(source: string): boolean {
  const candidate = maskLatexComments(source).trim();
  if (!candidate.startsWith("\\begin")) return false;

  const stack: string[] = [];
  let rootEnd = -1;
  ENVIRONMENT_TOKEN_PATTERN.lastIndex = 0;

  for (
    let match = ENVIRONMENT_TOKEN_PATTERN.exec(candidate);
    match;
    match = ENVIRONMENT_TOKEN_PATTERN.exec(candidate)
  ) {
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
