import type { Theme } from "./types/formula";
import { safeStorage } from "./runtime/safeStorage";

export const CUSTOM_THEME_STORAGE_KEY = "visualtex.custom-theme.v1";
const CUSTOM_THEME_CHANNEL = "visualtex-custom-theme";

export type ThemePaletteMode = "light" | "dark";

export interface ThemePaletteColors {
  accent: string;
  accentHover: string;
  accentSoft: string;
  background: string;
  elevated: string;
  surface: string;
  sunken: string;
  hover: string;
  active: string;
  foreground: string;
  textMuted: string;
  textFaint: string;
  border: string;
  borderStrong: string;
  formulaSurface: string;
  formulaPlaceholder: string;
  formulaPlaceholderSelected: string;
  formulaCaret: string;
  info: string;
  success: string;
  warning: string;
  danger: string;
  focusRing: string;
  syntaxCommand: string;
  syntaxKeyword: string;
  syntaxOperator: string;
  syntaxNumber: string;
  syntaxBracket: string;
  syntaxString: string;
  syntaxComment: string;
  syntaxVariable: string;
  syntaxFunction: string;
  syntaxError: string;
  toolbarCommon: string;
  toolbarStructure: string;
  toolbarCalculus: string;
  toolbarMatrix: string;
  toolbarRelation: string;
  toolbarGreek: string;
  toolbarArrow: string;
  toolbarPhysics: string;
  toolbarSet: string;
}

export interface ThemeDefinition {
  id: Theme;
  labelEn: string;
  labelZh: string;
  mode: ThemePaletteMode;
  swatches: readonly [string, string, string];
  colors?: ThemePaletteColors;
}

export interface CustomThemeState {
  version: 1;
  mode: ThemePaletteMode;
  colors: ThemePaletteColors;
}

const hexColorPattern = /^#[0-9a-f]{6}$/i;

function palette(
  mode: ThemePaletteMode,
  values: Partial<ThemePaletteColors> &
    Pick<ThemePaletteColors, "accent" | "background" | "surface" | "foreground">,
): ThemePaletteColors {
  const defaults: ThemePaletteColors =
    mode === "dark"
      ? {
          accent: "#72B7DD",
          accentHover: "#91CAE7",
          accentSoft: "#1D3440",
          background: "#16181B",
          elevated: "#1C1F23",
          surface: "#202328",
          sunken: "#131518",
          hover: "#2B3036",
          active: "#253B49",
          foreground: "#F2F4F6",
          textMuted: "#B8C0CA",
          textFaint: "#929DAA",
          border: "#353B43",
          borderStrong: "#4B535E",
          formulaSurface: "#202328",
          formulaPlaceholder: "#5F8FA8",
          formulaPlaceholderSelected: "#77A9C1",
          formulaCaret: "#72B7DD",
          info: "#72B7DD",
          success: "#62C99A",
          warning: "#E8B55A",
          danger: "#F08089",
          focusRing: "#78C7EF",
          syntaxCommand: "#A9DDF8",
          syntaxKeyword: "#A9DDF8",
          syntaxOperator: "#C8EAFF",
          syntaxNumber: "#F0B7D4",
          syntaxBracket: "#F0CF85",
          syntaxString: "#E7B08E",
          syntaxComment: "#9BA8B5",
          syntaxVariable: "#F2F4F6",
          syntaxFunction: "#B9E3F8",
          syntaxError: "#FF9AA4",
          toolbarCommon: "#A98F79",
          toolbarStructure: "#78A9DC",
          toolbarCalculus: "#D4AB60",
          toolbarMatrix: "#79BD91",
          toolbarRelation: "#AA8BD2",
          toolbarGreek: "#D28DA8",
          toolbarArrow: "#72B8BD",
          toolbarPhysics: "#8395D0",
          toolbarSet: "#75B39A",
        }
      : {
          accent: "#1F638E",
          accentHover: "#174F73",
          accentSoft: "#E5F0F6",
          background: "#F2F4F6",
          elevated: "#F7F8FA",
          surface: "#FFFFFF",
          sunken: "#E9EDF1",
          hover: "#E9EDF1",
          active: "#DCE8EF",
          foreground: "#1D232B",
          textMuted: "#4E5967",
          textFaint: "#687483",
          border: "#D5DAE0",
          borderStrong: "#B9C1CB",
          formulaSurface: "#FFFFFF",
          formulaPlaceholder: "#D9EDF9",
          formulaPlaceholderSelected: "#CFE8F7",
          formulaCaret: "#1F638E",
          info: "#1F638E",
          success: "#147554",
          warning: "#955400",
          danger: "#B53643",
          focusRing: "#1473A9",
          syntaxCommand: "#175F8F",
          syntaxKeyword: "#175F8F",
          syntaxOperator: "#376F91",
          syntaxNumber: "#9B3F70",
          syntaxBracket: "#8A5B16",
          syntaxString: "#8A4E2C",
          syntaxComment: "#687483",
          syntaxVariable: "#1D232B",
          syntaxFunction: "#285F7C",
          syntaxError: "#B53643",
          toolbarCommon: "#D5C5B4",
          toolbarStructure: "#C6DDF8",
          toolbarCalculus: "#F6DBA7",
          toolbarMatrix: "#C9E8D3",
          toolbarRelation: "#D9C9EF",
          toolbarGreek: "#F1C6D5",
          toolbarArrow: "#C2E6E8",
          toolbarPhysics: "#CAD2EF",
          toolbarSet: "#C8E4D9",
        };
  return { ...defaults, ...values };
}

const EXTRA_THEME_DEFINITIONS: readonly ThemeDefinition[] = [
  {
    id: "codex",
    labelEn: "Codex",
    labelZh: "Codex",
    mode: "light",
    swatches: ["#FFFFFF", "#1A1C1F", "#339CFF"],
    colors: palette("light", {
      accent: "#339CFF",
      accentHover: "#168CFF",
      accentSoft: "#EAF5FF",
      background: "#FFFFFF",
      elevated: "#F7F7F7",
      surface: "#FFFFFF",
      sunken: "#F2F2F2",
      hover: "#F5F5F5",
      active: "#EAF5FF",
      foreground: "#1A1C1F",
      textMuted: "#666666",
      textFaint: "#8C8C8C",
      border: "#E6E6E6",
      borderStrong: "#CCCCCC",
      formulaSurface: "#FFFFFF",
      formulaPlaceholder: "#D9EEFF",
      formulaPlaceholderSelected: "#B8DDFF",
      formulaCaret: "#339CFF",
      info: "#339CFF",
      success: "#18864B",
      warning: "#A96300",
      danger: "#C93C37",
      focusRing: "#339CFF",
      syntaxCommand: "#087EA4",
      syntaxKeyword: "#339CFF",
      syntaxOperator: "#666666",
      syntaxNumber: "#A96300",
      syntaxBracket: "#666666",
      syntaxString: "#18864B",
      syntaxComment: "#8C8C8C",
      syntaxVariable: "#1A1C1F",
      syntaxFunction: "#087EA4",
      syntaxError: "#C93C37",
      toolbarCommon: "#8C8C8C",
      toolbarStructure: "#339CFF",
      toolbarCalculus: "#A96300",
      toolbarMatrix: "#18864B",
      toolbarRelation: "#7457D9",
      toolbarGreek: "#C2477A",
      toolbarArrow: "#087EA4",
      toolbarPhysics: "#5267C9",
      toolbarSet: "#18864B",
    }),
  },
  {
    id: "notion",
    labelEn: "Notion",
    labelZh: "Notion",
    mode: "light",
    swatches: ["#FFFFFF", "#37352F", "#4981D2"],
    colors: palette("light", {
      accent: "#4981D2",
      accentHover: "#3A6FB9",
      accentSoft: "#EDF3FB",
      background: "#FFFFFF",
      elevated: "#F7F7F5",
      surface: "#FFFFFF",
      sunken: "#F1F1EF",
      hover: "#EFEFED",
      active: "#E7F3FB",
      foreground: "#37352F",
      textMuted: "#787774",
      textFaint: "#9B9A97",
      border: "#E3E2E0",
      borderStrong: "#D3D1CB",
      formulaSurface: "#FFFFFF",
      formulaPlaceholder: "#D3E5EF",
      formulaPlaceholderSelected: "#BFD9E7",
      formulaCaret: "#2383E2",
      info: "#337EA9",
      success: "#448361",
      warning: "#D9730D",
      danger: "#E03E3E",
      focusRing: "#2383E2",
      syntaxCommand: "#337EA9",
      syntaxKeyword: "#9065B0",
      syntaxOperator: "#9F6B53",
      syntaxNumber: "#D9730D",
      syntaxBracket: "#787774",
      syntaxString: "#448361",
      syntaxComment: "#9B9A97",
      syntaxVariable: "#37352F",
      syntaxFunction: "#337EA9",
      syntaxError: "#E03E3E",
      toolbarCommon: "#EDECE9",
      toolbarStructure: "#D3E5EF",
      toolbarCalculus: "#FDECC8",
      toolbarMatrix: "#DBEDDB",
      toolbarRelation: "#E8DEEE",
      toolbarGreek: "#F5E0E9",
      toolbarArrow: "#D3E5EF",
      toolbarPhysics: "#EEE0DA",
      toolbarSet: "#FADEC9",
    }),
  },
  {
    id: "one",
    labelEn: "One",
    labelZh: "One",
    mode: "light",
    swatches: ["#FAFAFA", "#383A42", "#586EF6"],
    colors: palette("light", {
      accent: "#586EF6",
      accentHover: "#465CE1",
      accentSoft: "#EEF0FF",
      background: "#FAFAFA",
      elevated: "#F6F6F6",
      surface: "#FAFAFA",
      sunken: "#F0F0F0",
      hover: "#F2F2F2",
      active: "#E8EBFF",
      foreground: "#383A42",
      textMuted: "#696C77",
      textFaint: "#A0A1A7",
      border: "#E5E5E6",
      borderStrong: "#D4D4D5",
      formulaSurface: "#FAFAFA",
      formulaPlaceholder: "#DDE3FF",
      formulaPlaceholderSelected: "#C6D0FF",
      formulaCaret: "#526FFF",
      info: "#4078F2",
      success: "#50A14F",
      warning: "#C18401",
      danger: "#E45649",
      focusRing: "#526FFF",
      syntaxCommand: "#0184BC",
      syntaxKeyword: "#A626A4",
      syntaxOperator: "#383A42",
      syntaxNumber: "#986801",
      syntaxBracket: "#383A42",
      syntaxString: "#50A14F",
      syntaxComment: "#A0A1A7",
      syntaxVariable: "#E45649",
      syntaxFunction: "#4078F2",
      syntaxError: "#CA1243",
      toolbarCommon: "#A0A1A7",
      toolbarStructure: "#4078F2",
      toolbarCalculus: "#C18401",
      toolbarMatrix: "#50A14F",
      toolbarRelation: "#A626A4",
      toolbarGreek: "#E45649",
      toolbarArrow: "#0184BC",
      toolbarPhysics: "#526FFF",
      toolbarSet: "#50A14F",
    }),
  },
  {
    id: "proof",
    labelEn: "Proof",
    labelZh: "Proof",
    mode: "light",
    swatches: ["#F5F3EE", "#2D312E", "#4B745F"],
    colors: palette("light", {
      accent: "#4B745F",
      accentHover: "#3D624F",
      accentSoft: "#E3EAE5",
      background: "#F5F3EE",
      elevated: "#F8F6F1",
      surface: "#F5F3EE",
      sunken: "#ECE9E2",
      hover: "#F0EEE8",
      active: "#DFE7E1",
      foreground: "#2D312E",
      textMuted: "#656A65",
      textFaint: "#92958F",
      border: "#D5D2CB",
      borderStrong: "#C3BEB4",
      formulaSurface: "#F5F3EE",
      formulaPlaceholder: "#DCE8E0",
      formulaPlaceholderSelected: "#C3D8CB",
      formulaCaret: "#4B745F",
      info: "#527589",
      success: "#4B745F",
      warning: "#A27434",
      danger: "#A6554B",
      focusRing: "#4B745F",
      syntaxCommand: "#527589",
      syntaxKeyword: "#6B6E4B",
      syntaxOperator: "#7A684F",
      syntaxNumber: "#A27434",
      syntaxBracket: "#7A746A",
      syntaxString: "#4B745F",
      syntaxComment: "#92958F",
      syntaxVariable: "#2D312E",
      syntaxFunction: "#527589",
      syntaxError: "#A6554B",
      toolbarCommon: "#D5D2CB",
      toolbarStructure: "#B8C8C0",
      toolbarCalculus: "#D8C7A8",
      toolbarMatrix: "#C4D3C8",
      toolbarRelation: "#D0CDC5",
      toolbarGreek: "#C8BDB4",
      toolbarArrow: "#B9CAC3",
      toolbarPhysics: "#BFC5C0",
      toolbarSet: "#B8C9BC",
    }),
  },
  {
    id: "raycast",
    labelEn: "Raycast",
    labelZh: "Raycast",
    mode: "light",
    swatches: ["#FFFFFF", "#1D1D1F", "#ED6E69"],
    colors: palette("light", {
      accent: "#ED6E69",
      accentHover: "#D85D59",
      accentSoft: "#FFF0F0",
      background: "#FFFFFF",
      elevated: "#F8F8F8",
      surface: "#FFFFFF",
      sunken: "#F2F2F2",
      hover: "#F5F5F5",
      active: "#FFF0F0",
      foreground: "#1D1D1F",
      textMuted: "#666666",
      textFaint: "#929292",
      border: "#E6E6E6",
      borderStrong: "#D0D0D0",
      formulaSurface: "#FFFFFF",
      formulaPlaceholder: "#FFE1E1",
      formulaPlaceholderSelected: "#FFCACA",
      formulaCaret: "#ED6E69",
      info: "#3478F6",
      success: "#34A853",
      warning: "#C67A00",
      danger: "#ED6E69",
      focusRing: "#ED6E69",
      syntaxCommand: "#3478F6",
      syntaxKeyword: "#AF52DE",
      syntaxOperator: "#666666",
      syntaxNumber: "#C67A00",
      syntaxBracket: "#666666",
      syntaxString: "#2E8B57",
      syntaxComment: "#929292",
      syntaxVariable: "#1D1D1F",
      syntaxFunction: "#3478F6",
      syntaxError: "#FF3B30",
      toolbarCommon: "#929292",
      toolbarStructure: "#3478F6",
      toolbarCalculus: "#FF9500",
      toolbarMatrix: "#34C759",
      toolbarRelation: "#AF52DE",
      toolbarGreek: "#FF2D55",
      toolbarArrow: "#55B3FF",
      toolbarPhysics: "#5856D6",
      toolbarSet: "#34A853",
    }),
  },
  {
    id: "rose-pine",
    labelEn: "Rose Pine",
    labelZh: "Rose Pine",
    mode: "light",
    swatches: ["#F9F4EE", "#575279", "#CB8681"],
    colors: palette("light", {
      accent: "#CB8681",
      accentHover: "#B4637A",
      accentSoft: "#F2E9E5",
      background: "#F9F4EE",
      elevated: "#FCF8F3",
      surface: "#F9F4EE",
      sunken: "#F0E8E2",
      hover: "#F4EDE8",
      active: "#DFDAD9",
      foreground: "#575279",
      textMuted: "#797593",
      textFaint: "#9893A5",
      border: "#DFDAD9",
      borderStrong: "#CECACD",
      formulaSurface: "#F9F4EE",
      formulaPlaceholder: "#F2E9E1",
      formulaPlaceholderSelected: "#DFDAD9",
      formulaCaret: "#CB8681",
      info: "#286983",
      success: "#56949F",
      warning: "#EA9D34",
      danger: "#B4637A",
      focusRing: "#CB8681",
      syntaxCommand: "#286983",
      syntaxKeyword: "#907AA9",
      syntaxOperator: "#D7827E",
      syntaxNumber: "#EA9D34",
      syntaxBracket: "#797593",
      syntaxString: "#56949F",
      syntaxComment: "#9893A5",
      syntaxVariable: "#575279",
      syntaxFunction: "#286983",
      syntaxError: "#B4637A",
      toolbarCommon: "#9893A5",
      toolbarStructure: "#286983",
      toolbarCalculus: "#EA9D34",
      toolbarMatrix: "#56949F",
      toolbarRelation: "#907AA9",
      toolbarGreek: "#B4637A",
      toolbarArrow: "#D7827E",
      toolbarPhysics: "#286983",
      toolbarSet: "#56949F",
    }),
  },
  {
    id: "solarized",
    labelEn: "Solarized",
    labelZh: "Solarized",
    mode: "light",
    swatches: ["#FCF6E5", "#657B83", "#AE8B2D"],
    colors: palette("light", {
      accent: "#AE8B2D",
      accentHover: "#8F6D00",
      accentSoft: "#F3E8B7",
      background: "#FCF6E5",
      elevated: "#F8F0DD",
      surface: "#FCF6E5",
      sunken: "#E8E1CE",
      hover: "#F4EEDC",
      active: "#EDE1AE",
      foreground: "#657B83",
      textMuted: "#839496",
      textFaint: "#93A1A1",
      border: "#DDD6C3",
      borderStrong: "#C9C1AC",
      formulaSurface: "#FCF6E5",
      formulaPlaceholder: "#EEE2A7",
      formulaPlaceholderSelected: "#DECC79",
      formulaCaret: "#AE8B2D",
      info: "#268BD2",
      success: "#859900",
      warning: "#CB4B16",
      danger: "#DC322F",
      focusRing: "#AE8B2D",
      syntaxCommand: "#268BD2",
      syntaxKeyword: "#859900",
      syntaxOperator: "#2AA198",
      syntaxNumber: "#6C71C4",
      syntaxBracket: "#B58900",
      syntaxString: "#2AA198",
      syntaxComment: "#93A1A1",
      syntaxVariable: "#657B83",
      syntaxFunction: "#268BD2",
      syntaxError: "#DC322F",
      toolbarCommon: "#93A1A1",
      toolbarStructure: "#268BD2",
      toolbarCalculus: "#B58900",
      toolbarMatrix: "#859900",
      toolbarRelation: "#6C71C4",
      toolbarGreek: "#D33682",
      toolbarArrow: "#2AA198",
      toolbarPhysics: "#CB4B16",
      toolbarSet: "#859900",
    }),
  },
  {
    id: "vercel",
    labelEn: "Vercel",
    labelZh: "Vercel",
    mode: "light",
    swatches: ["#FFFFFF", "#1A1A1A", "#2D69F6"],
    colors: palette("light", {
      accent: "#2D69F6",
      accentHover: "#1E58D8",
      accentSoft: "#EAF2FF",
      background: "#FAFAFA",
      elevated: "#FFFFFF",
      surface: "#FFFFFF",
      sunken: "#F2F2F2",
      hover: "#EBEBEB",
      active: "#E6E6E6",
      foreground: "#1A1A1A",
      textMuted: "#666666",
      textFaint: "#8F8F8F",
      border: "#EBEBEB",
      borderStrong: "#D8D8D8",
      formulaSurface: "#FFFFFF",
      formulaPlaceholder: "#EAF2FF",
      formulaPlaceholderSelected: "#C9DFFF",
      formulaCaret: "#2D69F6",
      info: "#0066FF",
      success: "#46A758",
      warning: "#E79D13",
      danger: "#E5484D",
      focusRing: "#2D69F6",
      syntaxCommand: "#0066FF",
      syntaxKeyword: "#8E4EC6",
      syntaxOperator: "#666666",
      syntaxNumber: "#E79D13",
      syntaxBracket: "#666666",
      syntaxString: "#46A758",
      syntaxComment: "#8F8F8F",
      syntaxVariable: "#1A1A1A",
      syntaxFunction: "#0066FF",
      syntaxError: "#E5484D",
      toolbarCommon: "#8F8F8F",
      toolbarStructure: "#0066FF",
      toolbarCalculus: "#E79D13",
      toolbarMatrix: "#46A758",
      toolbarRelation: "#8E4EC6",
      toolbarGreek: "#D6409F",
      toolbarArrow: "#12A594",
      toolbarPhysics: "#0066FF",
      toolbarSet: "#46A758",
    }),
  },
  {
    id: "vscode-plus",
    labelEn: "VS Code Plus",
    labelZh: "VS Code Plus",
    mode: "light",
    swatches: ["#FFFFFF", "#000000", "#3478C6"],
    colors: palette("light", {
      accent: "#3478C6",
      accentHover: "#2865AA",
      accentSoft: "#E5F3FF",
      background: "#F3F3F3",
      elevated: "#FFFFFF",
      surface: "#FFFFFF",
      sunken: "#F3F3F3",
      hover: "#E8E8E8",
      active: "#E4E6F1",
      foreground: "#000000",
      textMuted: "#616161",
      textFaint: "#767676",
      border: "#D4D4D4",
      borderStrong: "#CECECE",
      formulaSurface: "#FFFFFF",
      formulaPlaceholder: "#D6E9FF",
      formulaPlaceholderSelected: "#ADD6FF",
      formulaCaret: "#3478C6",
      info: "#007ACC",
      success: "#16825D",
      warning: "#A66A00",
      danger: "#CD3131",
      focusRing: "#007FD4",
      syntaxCommand: "#0451A5",
      syntaxKeyword: "#AF00DB",
      syntaxOperator: "#000000",
      syntaxNumber: "#098658",
      syntaxBracket: "#000000",
      syntaxString: "#A31515",
      syntaxComment: "#008000",
      syntaxVariable: "#001080",
      syntaxFunction: "#795E26",
      syntaxError: "#CD3131",
      toolbarCommon: "#767676",
      toolbarStructure: "#267F99",
      toolbarCalculus: "#795E26",
      toolbarMatrix: "#008000",
      toolbarRelation: "#AF00DB",
      toolbarGreek: "#A31515",
      toolbarArrow: "#0451A5",
      toolbarPhysics: "#0070C1",
      toolbarSet: "#16825D",
    }),
  },
  {
    id: "xcode",
    labelEn: "Xcode",
    labelZh: "Xcode",
    mode: "light",
    swatches: ["#FFFFFF", "#262626", "#0F0EF5"],
    colors: palette("light", {
      accent: "#0F0EF5",
      accentHover: "#0A0ACD",
      accentSoft: "#E5F1FF",
      background: "#F5F5F5",
      elevated: "#FAFAFA",
      surface: "#FFFFFF",
      sunken: "#ECECEC",
      hover: "#F0F0F0",
      active: "#DDEEFF",
      foreground: "#262626",
      textMuted: "#6C6C70",
      textFaint: "#8A99A6",
      border: "#D1D1D6",
      borderStrong: "#B9B9BE",
      formulaSurface: "#FFFFFF",
      formulaPlaceholder: "#D8E9FB",
      formulaPlaceholderSelected: "#B4D8FD",
      formulaCaret: "#0F0EF5",
      info: "#0F68A0",
      success: "#3E8087",
      warning: "#78492A",
      danger: "#D12F1B",
      focusRing: "#0F0EF5",
      syntaxCommand: "#0F68A0",
      syntaxKeyword: "#AD3DA4",
      syntaxOperator: "#262626",
      syntaxNumber: "#78492A",
      syntaxBracket: "#262626",
      syntaxString: "#D12F1B",
      syntaxComment: "#5D6C79",
      syntaxVariable: "#262626",
      syntaxFunction: "#0F68A0",
      syntaxError: "#D12F1B",
      toolbarCommon: "#8A99A6",
      toolbarStructure: "#0F68A0",
      toolbarCalculus: "#78492A",
      toolbarMatrix: "#3E8087",
      toolbarRelation: "#AD3DA4",
      toolbarGreek: "#D12F1B",
      toolbarArrow: "#0F68A0",
      toolbarPhysics: "#804FB8",
      toolbarSet: "#3E8087",
    }),
  }
];

export const THEME_DEFINITIONS: readonly ThemeDefinition[] = [
  {
    id: "light",
    labelEn: "Light",
    labelZh: "浅色",
    mode: "light",
    swatches: ["#F2F4F6", "#1D232B", "#1F638E"],
  },
  {
    id: "beige",
    labelEn: "Warm beige",
    labelZh: "暖米色",
    mode: "light",
    swatches: ["#E4D5BF", "#352C24", "#785536"],
  },
  {
    id: "dark",
    labelEn: "Dark",
    labelZh: "深色",
    mode: "dark",
    swatches: ["#16181B", "#F2F4F6", "#72B7DD"],
  },
  {
    id: "purple",
    labelEn: "Deep purple",
    labelZh: "深紫色",
    mode: "dark",
    swatches: ["#120E16", "#F5F0F8", "#BFA4EF"],
  },
  {
    id: "green",
    labelEn: "Deep green",
    labelZh: "深绿色",
    mode: "dark",
    swatches: ["#0D120F", "#EEF5F0", "#87CDAA"],
  },
  ...EXTRA_THEME_DEFINITIONS,
  {
    id: "custom",
    labelEn: "Custom",
    labelZh: "自定义",
    mode: "light",
    swatches: ["#FFFFFF", "#1A1C1F", "#339CFF"],
  },
] as const;

export const THEME_IDS = THEME_DEFINITIONS.map((item) => item.id) as readonly Theme[];

export function isTheme(value: unknown): value is Theme {
  return typeof value === "string" && THEME_IDS.includes(value as Theme);
}

export function getThemeDefinition(theme: Theme) {
  return THEME_DEFINITIONS.find((item) => item.id === theme) ?? THEME_DEFINITIONS[0];
}

function normalizeHex(value: unknown, fallback: string) {
  if (typeof value !== "string") return fallback;
  const trimmed = value.trim();
  if (hexColorPattern.test(trimmed)) return trimmed.toUpperCase();
  const shorthand = /^#([0-9a-f]{3})$/i.exec(trimmed);
  if (!shorthand) return fallback;
  return `#${shorthand[1]
    .split("")
    .map((part) => `${part}${part}`)
    .join("")}`.toUpperCase();
}

export function createDefaultCustomTheme(): CustomThemeState {
  const codex = EXTRA_THEME_DEFINITIONS.find((item) => item.id === "codex");
  if (!codex?.colors) throw new Error("Missing Codex theme palette");
  return {
    version: 1,
    mode: "light",
    colors: { ...codex.colors },
  };
}

export function normalizeCustomTheme(value: unknown): CustomThemeState {
  const fallback = createDefaultCustomTheme();
  if (!value || typeof value !== "object" || Array.isArray(value)) return fallback;
  const candidate = value as Partial<CustomThemeState> & {
    colors?: Partial<ThemePaletteColors>;
  };
  const sourceColors: Partial<ThemePaletteColors> =
    candidate.colors && typeof candidate.colors === "object"
      ? candidate.colors
      : {};
  const colors = Object.fromEntries(
    Object.entries(fallback.colors).map(([key, defaultValue]) => [
      key,
      normalizeHex(sourceColors[key as keyof ThemePaletteColors], defaultValue),
    ]),
  ) as unknown as ThemePaletteColors;
  return {
    version: 1,
    mode: candidate.mode === "dark" ? "dark" : "light",
    colors,
  };
}

export function readCustomTheme(): CustomThemeState {
  const raw = safeStorage.getItem(CUSTOM_THEME_STORAGE_KEY);
  if (!raw) return createDefaultCustomTheme();
  try {
    return normalizeCustomTheme(JSON.parse(raw));
  } catch {
    return createDefaultCustomTheme();
  }
}

export function saveCustomTheme(state: CustomThemeState) {
  const normalized = normalizeCustomTheme(state);
  safeStorage.setItem(CUSTOM_THEME_STORAGE_KEY, JSON.stringify(normalized));
}

function parseHex(hex: string) {
  const value = normalizeHex(hex, "#000000").slice(1);
  return {
    r: Number.parseInt(value.slice(0, 2), 16),
    g: Number.parseInt(value.slice(2, 4), 16),
    b: Number.parseInt(value.slice(4, 6), 16),
  };
}

function rgba(hex: string, alpha: number) {
  const { r, g, b } = parseHex(hex);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function foregroundOn(hex: string) {
  const { r, g, b } = parseHex(hex);
  const luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
  return luminance > 0.61 ? "#111318" : "#FFFFFF";
}

const inlineThemeProperties = [
  "color-scheme",
  "--bg",
  "--bg-elevated",
  "--surface",
  "--surface-primary",
  "--surface-secondary",
  "--surface-sunken",
  "--surface-hover",
  "--surface-active",
  "--bg-global",
  "--bg-base",
  "--bg-raised",
  "--bg-paper",
  "--bg-panel",
  "--bg-canvas",
  "--bg-sunken",
  "--bg-inset",
  "--bg-overlay",
  "--bg-modal",
  "--bg-disabled",
  "--bg-scrim",
  "--bg-hover",
  "--bg-active",
  "--bg-selected-subtle",
  "--bg-selected",
  "--bg-current-line",
  "--bg-drop-target",
  "--text",
  "--text-primary",
  "--text-secondary",
  "--text-muted",
  "--text-faint",
  "--text-muted-beige",
  "--text-muted-purple",
  "--text-muted-green",
  "--text-placeholder",
  "--text-placeholder-focus",
  "--text-placeholder-selected",
  "--text-disabled",
  "--text-on-accent",
  "--text-on-dark",
  "--icon-primary",
  "--icon-secondary",
  "--icon-muted",
  "--icon-disabled",
  "--border",
  "--border-subtle",
  "--border-default",
  "--border-strong",
  "--border-strong-beige",
  "--border-strong-purple",
  "--border-strong-green",
  "--border-focus",
  "--border-disabled",
  "--separator",
  "--accent",
  "--accent-primary",
  "--accent-hover",
  "--accent-hover-beige",
  "--accent-hover-purple",
  "--accent-hover-green",
  "--accent-active",
  "--accent-disabled",
  "--accent-soft",
  "--accent-soft-beige",
  "--accent-soft-purple",
  "--accent-soft-green",
  "--accent-soft-hover",
  "--accent-soft-active",
  "--accent-border",
  "--accent-foreground",
  "--focus-ring",
  "--focus-soft",
  "--interaction-primary",
  "--interaction-hover",
  "--interaction-active",
  "--interaction-soft",
  "--primary-foreground",
  "--formula-surface",
  "--formula-placeholder",
  "--formula-placeholder-selected",
  "--formula-caret",
  "--formula-selection",
  "--formula-selection-strong",
  "--formula-active-line",
  "--button-bg",
  "--button-bg-hover",
  "--button-bg-active",
  "--button-border",
  "--button-border-hover",
  "--button-text",
  "--input-bg",
  "--input-bg-hover",
  "--input-bg-focus",
  "--input-border",
  "--input-border-hover",
  "--input-border-focus",
  "--input-text",
  "--input-placeholder",
  "--tile-bg",
  "--tile-bg-hover",
  "--tile-bg-active",
  "--tile-bg-selected",
  "--tile-border",
  "--tile-border-hover",
  "--tile-border-selected",
  "--tile-text",
  "--tab-text",
  "--tab-text-active",
  "--tab-bg-hover",
  "--tab-bg-active",
  "--tab-indicator",
  "--menu-bg",
  "--tooltip-bg",
  "--scrollbar-track",
  "--scrollbar-thumb",
  "--scrollbar-thumb-hover",
  "--info",
  "--success",
  "--warning",
  "--warning-text",
  "--danger",
  "--danger-text",
  "--syntax-command",
  "--syntax-keyword",
  "--syntax-operator",
  "--syntax-number",
  "--syntax-bracket",
  "--syntax-string",
  "--syntax-comment",
  "--syntax-variable",
  "--syntax-function",
  "--syntax-error",
  "--shadow-sm",
  "--shadow-md",
  "--shadow-lg",
  "--toolbar-category-common",
  "--toolbar-category-structure",
  "--toolbar-category-calculus",
  "--toolbar-category-matrix",
  "--toolbar-category-relation",
  "--toolbar-category-greek",
  "--toolbar-category-arrow",
  "--toolbar-category-physics",
  "--toolbar-category-set",
] as const;

export function clearInlineThemePalette() {
  if (typeof document === "undefined") return;
  const root = document.documentElement;
  for (const property of inlineThemeProperties) root.style.removeProperty(property);
  delete root.dataset.themePalette;
}

function applyColors(colors: ThemePaletteColors, mode: ThemePaletteMode) {
  const root = document.documentElement;
  const primaryForeground = foregroundOn(colors.accent);
  const shadowTone = mode === "dark" ? "0, 0, 0" : "24, 31, 40";
  const values: Record<string, string> = {
    "color-scheme": mode,
    "--bg": colors.background,
    "--bg-elevated": colors.elevated,
    "--surface": colors.surface,
    "--surface-primary": colors.surface,
    "--surface-secondary": colors.elevated,
    "--surface-sunken": colors.sunken,
    "--surface-hover": colors.hover,
    "--surface-active": colors.active,
    "--bg-global": colors.background,
    "--bg-base": colors.background,
    "--bg-raised": colors.elevated,
    "--bg-paper": colors.surface,
    "--bg-panel": colors.surface,
    "--bg-canvas": colors.formulaSurface,
    "--bg-sunken": colors.sunken,
    "--bg-inset": colors.sunken,
    "--bg-overlay": colors.elevated,
    "--bg-modal": colors.surface,
    "--bg-disabled": colors.elevated,
    "--bg-scrim": rgba(colors.foreground, mode === "dark" ? 0.54 : 0.18),
    "--bg-hover": colors.hover,
    "--bg-active": colors.active,
    "--bg-selected-subtle": colors.accentSoft,
    "--bg-selected": colors.active,
    "--bg-current-line": colors.accentSoft,
    "--bg-drop-target": colors.active,
    "--text": colors.foreground,
    "--text-primary": colors.foreground,
    "--text-secondary": colors.textMuted,
    "--text-muted": colors.textMuted,
    "--text-faint": colors.textFaint,
    "--text-muted-beige": colors.textFaint,
    "--text-muted-purple": colors.textFaint,
    "--text-muted-green": colors.textFaint,
    "--text-placeholder": colors.textFaint,
    "--text-placeholder-focus": colors.textMuted,
    "--text-placeholder-selected": colors.formulaPlaceholderSelected,
    "--text-disabled": colors.textFaint,
    "--text-on-accent": primaryForeground,
    "--text-on-dark": mode === "dark" ? colors.foreground : "#FFFFFF",
    "--icon-primary": colors.foreground,
    "--icon-secondary": colors.textMuted,
    "--icon-muted": colors.textFaint,
    "--icon-disabled": colors.textFaint,
    "--border": colors.border,
    "--border-subtle": colors.border,
    "--border-default": colors.border,
    "--border-strong": colors.borderStrong,
    "--border-strong-beige": colors.borderStrong,
    "--border-strong-purple": colors.borderStrong,
    "--border-strong-green": colors.borderStrong,
    "--border-focus": colors.focusRing,
    "--border-disabled": colors.border,
    "--separator": rgba(colors.textMuted, 0.18),
    "--accent": colors.accent,
    "--accent-primary": colors.accent,
    "--accent-hover": colors.accentHover,
    "--accent-hover-beige": colors.accentHover,
    "--accent-hover-purple": colors.accentHover,
    "--accent-hover-green": colors.accentHover,
    "--accent-active": colors.accentHover,
    "--accent-disabled": colors.textFaint,
    "--accent-soft": colors.accentSoft,
    "--accent-soft-beige": colors.accentSoft,
    "--accent-soft-purple": colors.accentSoft,
    "--accent-soft-green": colors.accentSoft,
    "--accent-soft-hover": colors.hover,
    "--accent-soft-active": colors.active,
    "--accent-border": colors.borderStrong,
    "--accent-foreground": primaryForeground,
    "--focus-ring": colors.focusRing,
    "--focus-soft": rgba(colors.focusRing, 0.14),
    "--interaction-primary": colors.accent,
    "--interaction-hover": colors.accentHover,
    "--interaction-active": colors.accentHover,
    "--interaction-soft": colors.accentSoft,
    "--primary-foreground": primaryForeground,
    "--formula-surface": colors.formulaSurface,
    "--formula-placeholder": colors.formulaPlaceholder,
    "--formula-placeholder-selected": colors.formulaPlaceholderSelected,
    "--formula-caret": colors.formulaCaret,
    "--formula-selection": rgba(colors.accent, 0.22),
    "--formula-selection-strong": rgba(colors.accent, 0.36),
    "--formula-active-line": colors.accent,
    "--button-bg": colors.surface,
    "--button-bg-hover": colors.hover,
    "--button-bg-active": colors.active,
    "--button-border": colors.border,
    "--button-border-hover": colors.borderStrong,
    "--button-text": colors.foreground,
    "--input-bg": colors.formulaSurface,
    "--input-bg-hover": colors.elevated,
    "--input-bg-focus": colors.formulaSurface,
    "--input-border": colors.border,
    "--input-border-hover": colors.borderStrong,
    "--input-border-focus": colors.focusRing,
    "--input-text": colors.foreground,
    "--input-placeholder": colors.textFaint,
    "--tile-bg": colors.surface,
    "--tile-bg-hover": colors.hover,
    "--tile-bg-active": colors.active,
    "--tile-bg-selected": colors.accentSoft,
    "--tile-border": colors.border,
    "--tile-border-hover": colors.borderStrong,
    "--tile-border-selected": colors.accent,
    "--tile-text": colors.foreground,
    "--tab-text": colors.textMuted,
    "--tab-text-active": colors.accent,
    "--tab-bg-hover": colors.hover,
    "--tab-bg-active": colors.accentSoft,
    "--tab-indicator": colors.accent,
    "--menu-bg": colors.surface,
    "--tooltip-bg": colors.elevated,
    "--scrollbar-track": colors.background,
    "--scrollbar-thumb": colors.borderStrong,
    "--scrollbar-thumb-hover": colors.textFaint,
    "--info": colors.info,
    "--success": colors.success,
    "--warning": colors.warning,
    "--warning-text": colors.warning,
    "--danger": colors.danger,
    "--danger-text": colors.danger,
    "--syntax-command": colors.syntaxCommand,
    "--syntax-keyword": colors.syntaxKeyword,
    "--syntax-operator": colors.syntaxOperator,
    "--syntax-number": colors.syntaxNumber,
    "--syntax-bracket": colors.syntaxBracket,
    "--syntax-string": colors.syntaxString,
    "--syntax-comment": colors.syntaxComment,
    "--syntax-variable": colors.syntaxVariable,
    "--syntax-function": colors.syntaxFunction,
    "--syntax-error": colors.syntaxError,
    "--shadow-sm": `0 1px 2px rgba(${shadowTone}, ${mode === "dark" ? 0.22 : 0.07})`,
    "--shadow-md": `0 7px 20px rgba(${shadowTone}, ${mode === "dark" ? 0.32 : 0.13})`,
    "--shadow-lg": `0 16px 40px rgba(${shadowTone}, ${mode === "dark" ? 0.44 : 0.18})`,
    "--toolbar-category-common": colors.toolbarCommon,
    "--toolbar-category-structure": colors.toolbarStructure,
    "--toolbar-category-calculus": colors.toolbarCalculus,
    "--toolbar-category-matrix": colors.toolbarMatrix,
    "--toolbar-category-relation": colors.toolbarRelation,
    "--toolbar-category-greek": colors.toolbarGreek,
    "--toolbar-category-arrow": colors.toolbarArrow,
    "--toolbar-category-physics": colors.toolbarPhysics,
    "--toolbar-category-set": colors.toolbarSet,
  };
  for (const [property, value] of Object.entries(values)) {
    root.style.setProperty(property, value);
  }
}

export function applyThemePalette(theme: Theme) {
  if (typeof document === "undefined") return;
  clearInlineThemePalette();
  const root = document.documentElement;
  if (theme === "custom") {
    const custom = readCustomTheme();
    root.dataset.themePalette = "custom";
    applyColors(custom.colors, custom.mode);
    return;
  }
  const definition = getThemeDefinition(theme);
  if (definition.colors) {
    root.dataset.themePalette = theme;
    applyColors(definition.colors, definition.mode);
  }
}

export function publishCustomTheme(state: CustomThemeState) {
  const normalized = normalizeCustomTheme(state);
  saveCustomTheme(normalized);
  if (document.documentElement.dataset.theme === "custom") {
    applyThemePalette("custom");
  }
  if (typeof BroadcastChannel === "undefined") return;
  let channel: BroadcastChannel | null = null;
  try {
    channel = new BroadcastChannel(CUSTOM_THEME_CHANNEL);
    channel.postMessage(normalized);
  } catch {
    // The current window has already applied and persisted the custom theme.
  } finally {
    channel?.close();
  }
}

export function subscribeCustomTheme() {
  const applyIfCustom = () => {
    if (document.documentElement.dataset.theme === "custom") {
      applyThemePalette("custom");
    }
  };
  const handleStorage = (event: StorageEvent) => {
    if (event.key === CUSTOM_THEME_STORAGE_KEY) applyIfCustom();
  };
  window.addEventListener("storage", handleStorage);
  let channel: BroadcastChannel | null = null;
  if (typeof BroadcastChannel !== "undefined") {
    try {
      channel = new BroadcastChannel(CUSTOM_THEME_CHANNEL);
    } catch {
      channel = null;
    }
  }
  if (channel) channel.onmessage = applyIfCustom;
  return () => {
    window.removeEventListener("storage", handleStorage);
    channel?.close();
  };
}
