import type { Theme } from "./types/formula";
import {
  applyThemePalette,
  isTheme,
} from "./themeCustomization";

export function normalizeSynchronizedTheme(value: unknown): Theme {
  return isTheme(value) ? value : "light";
}

export function applyDocumentTheme(theme: Theme) {
  const normalized = normalizeSynchronizedTheme(theme);
  document.documentElement.dataset.theme = normalized;
  applyThemePalette(normalized);
}
