function isEscaped(source: string, index: number): boolean {
  let count = 0;
  while (index > 0 && source[--index] === "\\") count += 1;
  return count % 2 === 1;
}

/** TeX comments must be removed before a host flattens physical newlines. */
export function stripMathSourceComments(source: string): string {
  let result = "";
  for (let index = 0; index < source.length; index += 1) {
    if (source[index] === "%" && !isEscaped(source, index)) {
      while (index < source.length && !/[\r\n]/.test(source[index])) index += 1;
      if (source[index] === "\r" && source[index + 1] === "\n") index += 1;
      continue;
    }
    result += source[index];
  }
  return result;
}

/**
 * Editor rows already live in math mode. MathLive can retain redundant dollar
 * mode switches after paste; emitting them inside another math wrapper creates
 * invalid source. Only remove paired switches outside text arguments. Literal
 * currency and math deliberately embedded inside text are preserved.
 * This must never be applied to a document before its math spans are parsed.
 */
export function normalizeMathModeSource(value: string): string {
  const source = stripMathSourceComments(value);
  const switches: Array<{ start: number; end: number; token: string }> = [];
  for (let index = 0; index < source.length; index += 1) {
    if (isEscaped(source, index)) continue;
    if (source[index] === "\\") {
      const text = source.slice(index).match(/^\\(?:text(?:rm|normal|bf|it|sf|tt|up|sl|sc)?|mbox|hbox)\s*\{/);
      if (text) {
        let depth = 1;
        index += text[0].length;
        while (index < source.length && depth > 0) {
          if (!isEscaped(source, index)) {
            if (source[index] === "{") depth += 1;
            if (source[index] === "}") depth -= 1;
          }
          index += 1;
        }
        index -= 1;
        continue;
      }
    }
    if (source[index] !== "$") continue;
    const token = source.startsWith("$$", index) ? "$$" : "$";
    switches.push({ start: index, end: index + token.length, token });
    index += token.length - 1;
  }
  if (switches.length % 2 || switches.some((item, index) => index % 2 === 1 && item.token !== switches[index - 1].token)) {
    return source;
  }
  let result = "";
  let cursor = 0;
  for (const item of switches) {
    result += source.slice(cursor, item.start);
    cursor = item.end;
  }
  return result + source.slice(cursor);
}
