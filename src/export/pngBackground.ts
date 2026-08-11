export type PngExportBackground = "transparent" | `#${string}`;

export const DEFAULT_PNG_EXPORT_BACKGROUND: PngExportBackground = "transparent";

export function normalizePngExportBackground(
  value: unknown,
): PngExportBackground {
  if (typeof value !== "string") return DEFAULT_PNG_EXPORT_BACKGROUND;
  const normalized = value.trim().toLowerCase();
  if (!normalized || normalized === "transparent") {
    return DEFAULT_PNG_EXPORT_BACKGROUND;
  }
  if (/^#[0-9a-f]{6}$/.test(normalized)) {
    return normalized as PngExportBackground;
  }
  return DEFAULT_PNG_EXPORT_BACKGROUND;
}

export function pngExportBackgroundPickerValue(value: unknown) {
  const normalized = normalizePngExportBackground(value);
  return normalized === "transparent" ? "#ffffff" : normalized;
}
