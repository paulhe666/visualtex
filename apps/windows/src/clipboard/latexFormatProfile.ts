import type {
  DisplayLatexWrapper,
  FormulaDisplayStyle,
  LatexCodeFormat,
  LatexFormatProfile,
} from "../types/formula";

export const DEFAULT_LATEX_FORMAT_PROFILE: LatexFormatProfile = {
  inlineWrapper: "dollar",
  inlineTextPolicy: "outside-math",
  displayWrapper: "double-dollar",
  numbered: false,
  multilineEnvironment: "gather",
};

export function normalizeLatexFormatProfile(
  value: unknown,
  fallback: LatexFormatProfile = DEFAULT_LATEX_FORMAT_PROFILE,
): LatexFormatProfile {
  const candidate =
    value && typeof value === "object"
      ? (value as Partial<LatexFormatProfile>)
      : {};
  return {
    inlineWrapper:
      candidate.inlineWrapper === "paren" ? "paren" : fallback.inlineWrapper,
    inlineTextPolicy:
      candidate.inlineTextPolicy === "text-command" ||
      candidate.inlineTextPolicy === "outside-math"
        ? candidate.inlineTextPolicy
        : fallback.inlineTextPolicy,
    displayWrapper:
      candidate.displayWrapper === "bracket" ||
      candidate.displayWrapper === "equation" ||
      candidate.displayWrapper === "double-dollar"
        ? candidate.displayWrapper
        : fallback.displayWrapper,
    numbered:
      typeof candidate.numbered === "boolean"
        ? candidate.numbered
        : fallback.numbered,
    multilineEnvironment:
      candidate.multilineEnvironment === "align"
        ? "align"
        : candidate.multilineEnvironment === "gather"
          ? "gather"
          : fallback.multilineEnvironment,
  };
}

export function legacyCodeFormatToProfile(
  format: LatexCodeFormat | undefined,
): LatexFormatProfile {
  const profile = { ...DEFAULT_LATEX_FORMAT_PROFILE };
  switch (format) {
    case "inline-dollar":
      profile.inlineWrapper = "dollar";
      profile.inlineTextPolicy = "text-command";
      break;
    case "inline-text-double-dollar":
      profile.inlineWrapper = "dollar";
      profile.inlineTextPolicy = "outside-math";
      break;
    case "inline-paren":
      profile.inlineWrapper = "paren";
      profile.inlineTextPolicy = "text-command";
      break;
    case "display-bracket":
      profile.displayWrapper = "bracket";
      break;
    case "equation":
      profile.displayWrapper = "equation";
      profile.numbered = true;
      break;
    case "equation-star":
      profile.displayWrapper = "equation";
      profile.numbered = false;
      break;
    case "align":
      profile.multilineEnvironment = "align";
      profile.numbered = true;
      break;
    case "align-star":
    case "aligned":
    case "equation-split":
    case "equation-star-split":
      profile.multilineEnvironment = "align";
      profile.numbered = format === "equation-split";
      break;
    case "gather":
      profile.multilineEnvironment = "gather";
      profile.numbered = true;
      break;
    case "gather-star":
    case "multline":
    case "multline-star":
      profile.multilineEnvironment = "gather";
      profile.numbered = format === "multline";
      break;
    default:
      break;
  }
  return profile;
}

export function resolveFormulaDisplayStyle(
  style: FormulaDisplayStyle | undefined,
  profile: LatexFormatProfile,
): Exclude<FormulaDisplayStyle, "default"> {
  if (style && style !== "default") return style;
  if (profile.displayWrapper === "bracket") return "bracket";
  if (profile.displayWrapper === "equation") {
    return profile.numbered ? "equation" : "equation-star";
  }
  return "double-dollar";
}

export function displayStyleLabel(
  style: FormulaDisplayStyle | undefined,
  profile: LatexFormatProfile,
) {
  const resolved = resolveFormulaDisplayStyle(style, profile);
  switch (resolved) {
    case "bracket":
      return "\\[ ]";
    case "equation":
      return "eq";
    case "equation-star":
      return "eq*";
    default:
      return "$$";
  }
}

export function displayStyleWrapper(
  style: FormulaDisplayStyle | undefined,
  profile: LatexFormatProfile,
): DisplayLatexWrapper {
  const resolved = resolveFormulaDisplayStyle(style, profile);
  if (resolved === "bracket") return "bracket";
  if (resolved === "equation" || resolved === "equation-star") {
    return "equation";
  }
  return "double-dollar";
}
