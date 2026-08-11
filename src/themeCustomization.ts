import type { Theme } from "./types/formula";
import { safeStorage } from "./runtime/safeStorage";

export interface ThemePaletteColors {
  accent: string;
  background: string;
  surface: string;
  elevated: string;
  sunken: string;
  foreground: string;
  textMuted: string;
  border: string;
  formulaSurface: string;
  formulaPlaceholder: string;
  formulaCaret: string;
  toolbarStructure: string;
  toolbarCalculus: string;
  toolbarMatrix: string;
  toolbarGreek: string;
}

export interface CustomThemeState {
  version: 1;
  mode: "light" | "dark";
  colors: ThemePaletteColors;
}

export interface ThemeDefinition {
  id: Theme;
  labelZh: string;
  labelEn: string;
  mode: "light" | "dark";
  swatches: string[];
  colors: ThemePaletteColors;
}

const CUSTOM_THEME_STORAGE_KEY = "visualtex.custom-theme.v1";
const CUSTOM_THEME_EVENT = "visualtex-custom-theme-changed";
const CUSTOM_THEME_CHANNEL = "visualtex-custom-theme";

const light: ThemePaletteColors = {
  accent: "#456A55",
  background: "#F4F5F2",
  surface: "#FFFFFF",
  elevated: "#FAFBF9",
  sunken: "#ECEFEB",
  foreground: "#172019",
  textMuted: "#657069",
  border: "#D7DDD8",
  formulaSurface: "#FFFFFF",
  formulaPlaceholder: "#A8B3AA",
  formulaCaret: "#365845",
  toolbarStructure: "#6677A8",
  toolbarCalculus: "#4D8367",
  toolbarMatrix: "#8A6E9E",
  toolbarGreek: "#A06F4C",
};

const palette = (
  partial: Partial<ThemePaletteColors>,
): ThemePaletteColors => ({ ...light, ...partial });

export const THEME_DEFINITIONS: ThemeDefinition[] = [
  {
    id: "light",
    labelZh: "浅色",
    labelEn: "Light",
    mode: "light",
    swatches: ["#F4F5F2", "#FFFFFF", "#456A55"],
    colors: light,
  },
  {
    id: "beige",
    labelZh: "米色",
    labelEn: "Beige",
    mode: "light",
    swatches: ["#F3EBDD", "#FBF4E8", "#8A7354"],
    colors: palette({
      accent: "#8A7354",
      background: "#F3EBDD",
      surface: "#FBF4E8",
      elevated: "#F8F0E4",
      sunken: "#F1E5D2",
      foreground: "#30271E",
      textMuted: "#796B5A",
      border: "#DDCFB6",
      formulaSurface: "#FBF4E8",
      formulaPlaceholder: "#CCB994",
      formulaCaret: "#70573C",
    }),
  },
  {
    id: "dark",
    labelZh: "深色",
    labelEn: "Dark",
    mode: "dark",
    swatches: ["#161A17", "#202622", "#7FC89B"],
    colors: palette({
      accent: "#7FC89B",
      background: "#151916",
      surface: "#202521",
      elevated: "#272E29",
      sunken: "#101310",
      foreground: "#EDF3EF",
      textMuted: "#A5B1A9",
      border: "#39423C",
      formulaSurface: "#1C211D",
      formulaPlaceholder: "#68746C",
      formulaCaret: "#8AD4A5",
      toolbarStructure: "#8FA4E6",
      toolbarCalculus: "#7FC89B",
      toolbarMatrix: "#C69BDD",
      toolbarGreek: "#DAA27A",
    }),
  },
  {
    id: "purple",
    labelZh: "紫色",
    labelEn: "Purple",
    mode: "light",
    swatches: ["#F3EFF8", "#FFFFFF", "#73558F"],
    colors: palette({
      accent: "#73558F",
      background: "#F3EFF8",
      elevated: "#FAF8FC",
      sunken: "#E9E1F1",
      border: "#D8CBE4",
      formulaPlaceholder: "#B29EC4",
      formulaCaret: "#674A82",
    }),
  },
  {
    id: "green",
    labelZh: "深绿",
    labelEn: "Forest",
    mode: "dark",
    swatches: ["#102019", "#183126", "#72C696"],
    colors: palette({
      accent: "#72C696",
      background: "#102019",
      surface: "#183126",
      elevated: "#1E3A2D",
      sunken: "#0B1711",
      foreground: "#E8F4ED",
      textMuted: "#A5BEAE",
      border: "#315642",
      formulaSurface: "#14291F",
      formulaPlaceholder: "#688675",
      formulaCaret: "#83D5A5",
    }),
  },
  {
    id: "codex",
    labelZh: "Codex",
    labelEn: "Codex",
    mode: "dark",
    swatches: ["#0D0D0D", "#171717", "#FAFAFA"],
    colors: palette({
      accent: "#ECECEC",
      background: "#0D0D0D",
      surface: "#171717",
      elevated: "#1F1F1F",
      sunken: "#080808",
      foreground: "#F2F2F2",
      textMuted: "#A6A6A6",
      border: "#323232",
      formulaSurface: "#141414",
      formulaPlaceholder: "#707070",
      formulaCaret: "#FFFFFF",
    }),
  },
  {
    id: "notion",
    labelZh: "Notion",
    labelEn: "Notion",
    mode: "light",
    swatches: ["#F7F6F3", "#FFFFFF", "#37352F"],
    colors: palette({
      accent: "#37352F",
      background: "#F7F6F3",
      surface: "#FFFFFF",
      elevated: "#FBFAF8",
      sunken: "#EEEDE9",
      foreground: "#37352F",
      textMuted: "#787774",
      border: "#E3E2DE",
      formulaPlaceholder: "#AAA9A5",
      formulaCaret: "#37352F",
    }),
  },
  {
    id: "one",
    labelZh: "One Dark",
    labelEn: "One Dark",
    mode: "dark",
    swatches: ["#282C34", "#21252B", "#61AFEF"],
    colors: palette({
      accent: "#61AFEF",
      background: "#21252B",
      surface: "#282C34",
      elevated: "#2F343D",
      sunken: "#1B1E23",
      foreground: "#ABB2BF",
      textMuted: "#7F848E",
      border: "#3B4048",
      formulaSurface: "#252931",
      formulaPlaceholder: "#626A78",
      formulaCaret: "#61AFEF",
      toolbarCalculus: "#98C379",
      toolbarMatrix: "#C678DD",
      toolbarGreek: "#E5C07B",
    }),
  },
  {
    id: "proof",
    labelZh: "Proof",
    labelEn: "Proof",
    mode: "light",
    swatches: ["#FAF8F2", "#FFFEFB", "#263C69"],
    colors: palette({
      accent: "#263C69",
      background: "#FAF8F2",
      surface: "#FFFEFB",
      elevated: "#F6F2E8",
      sunken: "#EFEADD",
      foreground: "#22252A",
      textMuted: "#6E706E",
      border: "#DCD5C7",
      formulaPlaceholder: "#AAA38F",
      formulaCaret: "#263C69",
    }),
  },
  {
    id: "raycast",
    labelZh: "Raycast",
    labelEn: "Raycast",
    mode: "dark",
    swatches: ["#161616", "#202020", "#FF6363"],
    colors: palette({
      accent: "#FF6363",
      background: "#121212",
      surface: "#1D1D1D",
      elevated: "#252525",
      sunken: "#0B0B0B",
      foreground: "#F5F5F5",
      textMuted: "#A0A0A0",
      border: "#353535",
      formulaSurface: "#181818",
      formulaPlaceholder: "#6B6B6B",
      formulaCaret: "#FF7272",
    }),
  },
  {
    id: "rose-pine",
    labelZh: "Rosé Pine",
    labelEn: "Rosé Pine",
    mode: "dark",
    swatches: ["#191724", "#1F1D2E", "#C4A7E7"],
    colors: palette({
      accent: "#C4A7E7",
      background: "#191724",
      surface: "#1F1D2E",
      elevated: "#26233A",
      sunken: "#13111B",
      foreground: "#E0DEF4",
      textMuted: "#908CAA",
      border: "#393552",
      formulaSurface: "#1D1A2B",
      formulaPlaceholder: "#6E6A86",
      formulaCaret: "#EBBCBA",
      toolbarCalculus: "#9CCFD8",
      toolbarMatrix: "#C4A7E7",
      toolbarGreek: "#F6C177",
    }),
  },
  {
    id: "solarized",
    labelZh: "Solarized",
    labelEn: "Solarized",
    mode: "light",
    swatches: ["#FDF6E3", "#EEE8D5", "#268BD2"],
    colors: palette({
      accent: "#268BD2",
      background: "#FDF6E3",
      surface: "#FFF9E9",
      elevated: "#FAF1DB",
      sunken: "#EEE8D5",
      foreground: "#586E75",
      textMuted: "#839496",
      border: "#D9D2BE",
      formulaSurface: "#FFF9E9",
      formulaPlaceholder: "#A9A38F",
      formulaCaret: "#268BD2",
      toolbarCalculus: "#859900",
      toolbarMatrix: "#6C71C4",
      toolbarGreek: "#B58900",
    }),
  },
  {
    id: "vercel",
    labelZh: "Vercel",
    labelEn: "Vercel",
    mode: "light",
    swatches: ["#FAFAFA", "#FFFFFF", "#000000"],
    colors: palette({
      accent: "#000000",
      background: "#FAFAFA",
      surface: "#FFFFFF",
      elevated: "#F7F7F7",
      sunken: "#EFEFEF",
      foreground: "#171717",
      textMuted: "#666666",
      border: "#E5E5E5",
      formulaPlaceholder: "#A3A3A3",
      formulaCaret: "#000000",
    }),
  },
  {
    id: "vscode-plus",
    labelZh: "VS Code+",
    labelEn: "VS Code+",
    mode: "dark",
    swatches: ["#1E1E1E", "#252526", "#007ACC"],
    colors: palette({
      accent: "#3794FF",
      background: "#1E1E1E",
      surface: "#252526",
      elevated: "#2D2D30",
      sunken: "#181818",
      foreground: "#D4D4D4",
      textMuted: "#9D9D9D",
      border: "#3E3E42",
      formulaSurface: "#202020",
      formulaPlaceholder: "#707070",
      formulaCaret: "#AEAFAD",
      toolbarCalculus: "#4EC9B0",
      toolbarMatrix: "#C586C0",
      toolbarGreek: "#DCDCAA",
    }),
  },
  {
    id: "xcode",
    labelZh: "Xcode",
    labelEn: "Xcode",
    mode: "light",
    swatches: ["#F5F5F5", "#FFFFFF", "#147EFB"],
    colors: palette({
      accent: "#147EFB",
      background: "#F5F5F5",
      surface: "#FFFFFF",
      elevated: "#FAFAFA",
      sunken: "#ECECEC",
      foreground: "#1F1F1F",
      textMuted: "#707070",
      border: "#D9D9D9",
      formulaPlaceholder: "#A0A0A0",
      formulaCaret: "#147EFB",
    }),
  },
  {
    id: "custom",
    labelZh: "自定义",
    labelEn: "Custom",
    mode: "light",
    swatches: ["#F4F5F2", "#FFFFFF", "#456A55"],
    colors: light,
  },
];

const themeIds = new Set<Theme>(THEME_DEFINITIONS.map((item) => item.id));

export function isTheme(value: unknown): value is Theme {
  return typeof value === "string" && themeIds.has(value as Theme);
}

export function getThemeDefinition(theme: Theme): ThemeDefinition {
  return (
    THEME_DEFINITIONS.find((definition) => definition.id === theme) ??
    THEME_DEFINITIONS[0]
  );
}

export function createDefaultCustomTheme(): CustomThemeState {
  return {
    version: 1,
    mode: "light",
    colors: { ...light },
  };
}

function normalizeHex(value: unknown, fallback: string) {
  if (typeof value !== "string") return fallback;
  const trimmed = value.trim();
  const match = /^#([0-9a-f]{6})$/i.exec(trimmed);
  return match ? `#${match[1].toUpperCase()}` : fallback;
}

function normalizeCustomTheme(value: unknown): CustomThemeState {
  if (!value || typeof value !== "object") return createDefaultCustomTheme();
  const candidate = value as Partial<CustomThemeState> & {
    colors?: Partial<ThemePaletteColors>;
  };
  const defaults = createDefaultCustomTheme();
  const colors = Object.fromEntries(
    (Object.keys(defaults.colors) as Array<keyof ThemePaletteColors>).map((key) => [
      key,
      normalizeHex(candidate.colors?.[key], defaults.colors[key]),
    ]),
  ) as unknown as ThemePaletteColors;
  return {
    version: 1,
    mode: candidate.mode === "dark" ? "dark" : "light",
    colors,
  };
}

export function readCustomTheme(): CustomThemeState {
  try {
    const raw = safeStorage.getItem(CUSTOM_THEME_STORAGE_KEY);
    return raw ? normalizeCustomTheme(JSON.parse(raw)) : createDefaultCustomTheme();
  } catch {
    return createDefaultCustomTheme();
  }
}

function dispatchCustomThemeChanged() {
  if (typeof window === "undefined") return;
  window.dispatchEvent(new CustomEvent(CUSTOM_THEME_EVENT));
  if (typeof BroadcastChannel !== "undefined") {
    const channel = new BroadcastChannel(CUSTOM_THEME_CHANNEL);
    channel.postMessage(Date.now());
    channel.close();
  }
}

export function publishCustomTheme(value: CustomThemeState) {
  const normalized = normalizeCustomTheme(value);
  safeStorage.setItem(CUSTOM_THEME_STORAGE_KEY, JSON.stringify(normalized));
  if (typeof document !== "undefined") {
    applyThemePalette("custom");
  }
  dispatchCustomThemeChanged();
  return normalized;
}

const cssVariables: Record<keyof ThemePaletteColors, string[]> = {
  accent: ["--accent", "--accent-primary", "--focus-ring"],
  background: ["--background", "--bg-app", "--bg-canvas"],
  surface: ["--surface", "--surface-primary", "--bg-paper", "--button-bg", "--input-bg"],
  elevated: ["--bg-elevated", "--bg-raised", "--surface-hover", "--button-bg-hover"],
  sunken: ["--bg-sunken", "--surface-sunken"],
  foreground: ["--foreground", "--text", "--text-primary", "--button-text", "--input-text"],
  textMuted: ["--text-muted", "--text-secondary", "--text-faint"],
  border: ["--border", "--border-default", "--border-subtle", "--button-border", "--input-border"],
  formulaSurface: ["--formula-surface", "--formula-background"],
  formulaPlaceholder: ["--formula-placeholder"],
  formulaCaret: ["--formula-caret"],
  toolbarStructure: ["--toolbar-structure"],
  toolbarCalculus: ["--toolbar-calculus"],
  toolbarMatrix: ["--toolbar-matrix"],
  toolbarGreek: ["--toolbar-greek"],
};

export function applyThemePalette(theme: Theme) {
  if (typeof document === "undefined") return;
  const definition = getThemeDefinition(theme);
  const custom = theme === "custom" ? readCustomTheme() : null;
  const colors = custom?.colors ?? definition.colors;
  const mode = custom?.mode ?? definition.mode;
  const root = document.documentElement;
  root.dataset.visualtexThemeMode = mode;
  root.style.colorScheme = mode;
  for (const key of Object.keys(cssVariables) as Array<keyof ThemePaletteColors>) {
    for (const cssVariable of cssVariables[key]) {
      root.style.setProperty(cssVariable, colors[key]);
    }
  }
  root.style.setProperty(
    "--accent-soft",
    `color-mix(in srgb, ${colors.accent} 13%, transparent)`,
  );
  root.style.setProperty(
    "--focus-soft",
    `color-mix(in srgb, ${colors.accent} 18%, transparent)`,
  );
}

export function subscribeCustomTheme(listener?: () => void): () => void {
  if (typeof window === "undefined") return () => undefined;
  const handler = () => listener?.();
  window.addEventListener(CUSTOM_THEME_EVENT, handler);
  const channel =
    typeof BroadcastChannel === "undefined"
      ? null
      : new BroadcastChannel(CUSTOM_THEME_CHANNEL);
  if (channel) channel.onmessage = handler;
  return () => {
    window.removeEventListener(CUSTOM_THEME_EVENT, handler);
    channel?.close();
  };
}
