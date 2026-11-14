export function normalizeCustomFormulaColor(value: unknown) {
  if (typeof value !== "string") return null;
  const normalized = value.trim().toLowerCase();
  const longHex = normalized.match(/^#([0-9a-f]{6})$/);
  if (longHex) return `#${longHex[1]}`;

  const shortHex = normalized.match(/^#([0-9a-f]{3})$/);
  if (shortHex) {
    return `#${Array.from(shortHex[1], (digit) => `${digit}${digit}`).join("")}`;
  }

  const rgb = normalized.match(
    /^rgba?\(\s*([+-]?(?:\d+\.?\d*|\.\d+)%?)\s*(?:,\s*|\s+)([+-]?(?:\d+\.?\d*|\.\d+)%?)\s*(?:,\s*|\s+)([+-]?(?:\d+\.?\d*|\.\d+)%?)(?:\s*(?:,|\/)\s*[+-]?(?:\d+\.?\d*|\.\d+)%?)?\s*\)$/,
  );
  if (!rgb) return null;

  const channel = (component: string) => {
    const percentage = component.endsWith("%");
    const numeric = Number.parseFloat(component);
    if (!Number.isFinite(numeric)) return null;
    return Math.round(
      Math.min(255, Math.max(0, percentage ? (numeric / 100) * 255 : numeric)),
    );
  };
  const channels = rgb.slice(1, 4).map(channel);
  if (channels.some((component) => component === null)) return null;
  return `#${channels
    .map((component) => component!.toString(16).padStart(2, "0"))
    .join("")}`;
}

export function isSafeFormulaStyleColor(value: unknown) {
  return (
    typeof value === "string" &&
    (/^#[0-9a-f]{6}$/i.test(value) ||
      [
        "red",
        "orange",
        "yellow",
        "lime",
        "green",
        "teal",
        "cyan",
        "blue",
        "indigo",
        "purple",
        "magenta",
        "black",
        "dark-grey",
        "grey",
        "light-grey",
        "white",
      ].includes(value))
  );
}
