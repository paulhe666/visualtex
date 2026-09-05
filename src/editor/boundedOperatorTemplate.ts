function readBracedScriptArgument(
  source: string,
  marker: "_" | "^",
  fromIndex: number,
) {
  const markerIndex = source.indexOf(marker, fromIndex);
  if (markerIndex < 0) return null;
  let cursor = markerIndex + 1;
  while (/\s/.test(source[cursor] ?? "")) cursor += 1;
  if (source[cursor] !== "{") return null;

  const contentStart = cursor + 1;
  let depth = 1;
  for (cursor = contentStart; cursor < source.length; cursor += 1) {
    if (source[cursor] === "{") depth += 1;
    else if (source[cursor] === "}") depth -= 1;
    if (depth === 0) {
      return {
        content: source.slice(contentStart, cursor),
        end: cursor + 1,
      };
    }
  }
  return null;
}

export function hasBoundedOperatorPlaceholderOrder(insertionTemplate: string) {
  const source = insertionTemplate.trim();
  const operator = source.match(/^\\[A-Za-z]+/);
  if (!operator) return false;

  let cursor = operator[0].length;
  while (/\s/.test(source[cursor] ?? "")) cursor += 1;
  if (source.startsWith("\\limits", cursor)) {
    cursor += "\\limits".length;
    while (/\s/.test(source[cursor] ?? "")) cursor += 1;
  }
  if (source[cursor] !== "_") return false;

  const lower = readBracedScriptArgument(source, "_", cursor);
  if (!lower || !lower.content.includes("\\placeholder{}")) return false;
  const upper = readBracedScriptArgument(source, "^", lower.end);
  return Boolean(upper?.content.includes("\\placeholder{}"));
}
