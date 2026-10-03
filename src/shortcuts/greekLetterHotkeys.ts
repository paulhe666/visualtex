import type { LatexCommand } from "../types/command";

type GreekShortcutDefinition = {
  id: string;
  latex: string;
  labelZh: string;
  labelEn: string;
};

type GreekShortcutEvent = Pick<
  KeyboardEvent,
  | "code"
  | "key"
  | "ctrlKey"
  | "altKey"
  | "shiftKey"
  | "metaKey"
  | "isComposing"
  | "repeat"
>;

const lowerGreekByCode: Record<string, GreekShortcutDefinition> = {
  KeyA: { id: "alpha", latex: "\\alpha", labelZh: "阿尔法", labelEn: "Alpha" },
  KeyB: { id: "beta", latex: "\\beta", labelZh: "贝塔", labelEn: "Beta" },
  KeyG: { id: "gamma", latex: "\\gamma", labelZh: "伽马", labelEn: "Gamma" },
  KeyD: { id: "delta", latex: "\\delta", labelZh: "德尔塔", labelEn: "Delta" },
  KeyE: { id: "epsilon", latex: "\\epsilon", labelZh: "艾普西隆", labelEn: "Epsilon" },
  KeyZ: { id: "zeta", latex: "\\zeta", labelZh: "泽塔", labelEn: "Zeta" },
  KeyH: { id: "eta", latex: "\\eta", labelZh: "伊塔", labelEn: "Eta" },
  KeyQ: { id: "theta", latex: "\\theta", labelZh: "西塔", labelEn: "Theta" },
  KeyI: { id: "iota", latex: "\\iota", labelZh: "约塔", labelEn: "Iota" },
  KeyK: { id: "kappa", latex: "\\kappa", labelZh: "卡帕", labelEn: "Kappa" },
  KeyL: { id: "lambda", latex: "\\lambda", labelZh: "拉姆达", labelEn: "Lambda" },
  KeyM: { id: "mu", latex: "\\mu", labelZh: "缪", labelEn: "Mu" },
  KeyN: { id: "nu", latex: "\\nu", labelZh: "纽", labelEn: "Nu" },
  KeyX: { id: "xi", latex: "\\xi", labelZh: "克西", labelEn: "Xi" },
  // Standard LaTeX has no \\omicron command because lowercase omicron is
  // visually identical to a Latin o. Use the conventional LaTeX spelling.
  KeyO: { id: "omicron", latex: "o", labelZh: "奥密克戎", labelEn: "Omicron" },
  KeyP: { id: "pi", latex: "\\pi", labelZh: "派", labelEn: "Pi" },
  KeyR: { id: "rho", latex: "\\rho", labelZh: "柔", labelEn: "Rho" },
  KeyS: { id: "sigma", latex: "\\sigma", labelZh: "西格玛", labelEn: "Sigma" },
  KeyT: { id: "tau", latex: "\\tau", labelZh: "陶", labelEn: "Tau" },
  KeyU: { id: "upsilon", latex: "\\upsilon", labelZh: "宇普西隆", labelEn: "Upsilon" },
  KeyF: { id: "phi", latex: "\\phi", labelZh: "斐", labelEn: "Phi" },
  KeyC: { id: "chi", latex: "\\chi", labelZh: "希", labelEn: "Chi" },
  KeyY: { id: "psi", latex: "\\psi", labelZh: "普赛", labelEn: "Psi" },
  KeyW: { id: "omega", latex: "\\omega", labelZh: "欧米伽", labelEn: "Omega" },
};

const upperGreekByCode: Record<string, GreekShortcutDefinition> = {
  KeyA: { id: "Alpha", latex: "A", labelZh: "大写阿尔法", labelEn: "Capital alpha" },
  KeyB: { id: "Beta", latex: "B", labelZh: "大写贝塔", labelEn: "Capital beta" },
  KeyG: { id: "Gamma", latex: "\\Gamma", labelZh: "大写伽马", labelEn: "Capital gamma" },
  KeyD: { id: "Delta", latex: "\\Delta", labelZh: "大写德尔塔", labelEn: "Capital delta" },
  KeyE: { id: "Epsilon", latex: "E", labelZh: "大写艾普西隆", labelEn: "Capital epsilon" },
  KeyZ: { id: "Zeta", latex: "Z", labelZh: "大写泽塔", labelEn: "Capital zeta" },
  KeyH: { id: "Eta", latex: "H", labelZh: "大写伊塔", labelEn: "Capital eta" },
  KeyQ: { id: "Theta", latex: "\\Theta", labelZh: "大写西塔", labelEn: "Capital theta" },
  KeyI: { id: "Iota", latex: "I", labelZh: "大写约塔", labelEn: "Capital iota" },
  KeyK: { id: "Kappa", latex: "K", labelZh: "大写卡帕", labelEn: "Capital kappa" },
  KeyL: { id: "Lambda", latex: "\\Lambda", labelZh: "大写拉姆达", labelEn: "Capital lambda" },
  KeyM: { id: "Mu", latex: "M", labelZh: "大写缪", labelEn: "Capital mu" },
  KeyN: { id: "Nu", latex: "N", labelZh: "大写纽", labelEn: "Capital nu" },
  KeyX: { id: "Xi", latex: "\\Xi", labelZh: "大写克西", labelEn: "Capital xi" },
  KeyO: { id: "Omicron", latex: "O", labelZh: "大写奥密克戎", labelEn: "Capital omicron" },
  KeyP: { id: "Pi", latex: "\\Pi", labelZh: "大写派", labelEn: "Capital pi" },
  // Greek capital rho and chi are conventionally represented by the same
  // glyphs as Latin P and X in LaTeX math mode.
  KeyR: { id: "Rho", latex: "P", labelZh: "大写柔", labelEn: "Capital rho" },
  KeyS: { id: "Sigma", latex: "\\Sigma", labelZh: "大写西格玛", labelEn: "Capital sigma" },
  KeyT: { id: "Tau", latex: "T", labelZh: "大写陶", labelEn: "Capital tau" },
  KeyU: { id: "Upsilon", latex: "\\Upsilon", labelZh: "大写宇普西隆", labelEn: "Capital upsilon" },
  KeyF: { id: "Phi", latex: "\\Phi", labelZh: "大写斐", labelEn: "Capital phi" },
  KeyC: { id: "Chi", latex: "X", labelZh: "大写希", labelEn: "Capital chi" },
  KeyY: { id: "Psi", latex: "\\Psi", labelZh: "大写普赛", labelEn: "Capital psi" },
  KeyW: { id: "Omega", latex: "\\Omega", labelZh: "大写欧米伽", labelEn: "Capital omega" },
};

function toLatexCommand(definition: GreekShortcutDefinition): LatexCommand {
  return {
    id: `greek-hotkey-${definition.id}`,
    command: definition.latex,
    insertTemplate: definition.latex,
    previewLatex: definition.latex,
    labelZh: definition.labelZh,
    labelEn: definition.labelEn,
    aliases: [],
    keywords: ["希腊字母", "Greek"],
    category: "greek",
    defaultPriority: 0,
    supportedInMathMode: true,
  };
}

export function isGreekLetterHotkeyPrefix(event: GreekShortcutEvent) {
  return (
    !event.repeat &&
    !event.isComposing &&
    event.code === "KeyG" &&
    event.metaKey &&
    !event.ctrlKey &&
    !event.altKey &&
    !event.shiftKey
  );
}

export function greekLetterHotkeyCommandFromEvent(
  event: GreekShortcutEvent,
): LatexCommand | null {
  if (
    event.repeat ||
    event.isComposing ||
    event.ctrlKey ||
    event.altKey ||
    event.metaKey ||
    event.key === "Process"
  ) {
    return null;
  }
  const definition = event.shiftKey
    ? upperGreekByCode[event.code]
    : lowerGreekByCode[event.code];
  return definition ? toLatexCommand(definition) : null;
}

export const greekLetterHotkeyLowercaseCodes = Object.keys(lowerGreekByCode);
