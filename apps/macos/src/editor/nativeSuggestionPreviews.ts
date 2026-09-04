import { compatibilityCommands } from "../autocomplete/compatibilityCommands";
import { EXTENDED_INTEGRAL_COMMANDS } from "../math/extendedIntegralCompatibility";

export type NativeSuggestionPreviewKind =
  | "arguments"
  | "delimiter"
  | "state"
  | "alias"
  | "spacing"
  | "operator"
  | "fallback";

export interface NativeSuggestionPreview {
  latex: string;
  kind: NativeSuggestionPreviewKind;
}

const previewRegistry = new Map<string, NativeSuggestionPreview>();

function registerPreviews(
  kind: NativeSuggestionPreviewKind,
  entries: ReadonlyArray<readonly [command: string, latex: string]>,
) {
  for (const [command, latex] of entries) {
    previewRegistry.set(command, { latex, kind });
  }
}

registerPreviews("arguments", [
  ["\\sqrt", "\\sqrt{x}"],
  ["\\frac", "\\frac{a}{b}"],
  ["\\dfrac", "\\dfrac{a}{b}"],
  ["\\tfrac", "\\tfrac{a}{b}"],
  ["\\binom", "\\binom{n}{k}"],
  ["\\dbinom", "\\dbinom{n}{k}"],
  ["\\tbinom", "\\tbinom{n}{k}"],
  ["\\cfrac", "\\cfrac{a}{b}"],
  ["\\pdiff", "\\pdiff{f}{x}"],
  ["\\nicefrac", "\\nicefrac{1}{2}"],
  ["\\acute", "\\acute{x}"],
  ["\\grave", "\\grave{x}"],
  ["\\dot", "\\dot{x}"],
  ["\\ddot", "\\ddot{x}"],
  ["\\dddot", "\\dddot{x}"],
  ["\\ddddot", "\\ddddot{x}"],
  ["\\tilde", "\\tilde{x}"],
  ["\\bar", "\\bar{x}"],
  ["\\breve", "\\breve{x}"],
  ["\\check", "\\check{x}"],
  ["\\hat", "\\hat{x}"],
  ["\\vec", "\\vec{x}"],
  ["\\mathring", "\\mathring{x}"],
  ["\\widehat", "\\widehat{ABC}"],
  ["\\widecheck", "\\widecheck{ABC}"],
  ["\\widetilde", "\\widetilde{ABC}"],
  ["\\utilde", "\\utilde{ABC}"],
  ["\\overarc", "\\overarc{AB}"],
  ["\\overline", "\\overline{AB}"],
  ["\\overbrace", "\\overbrace{a+b}"],
  ["\\overgroup", "\\overgroup{AB}"],
  ["\\overparen", "\\overparen{AB}"],
  ["\\wideparen", "\\wideparen{ABC}"],
  ["\\overleftarrow", "\\overleftarrow{AB}"],
  ["\\overrightarrow", "\\overrightarrow{AB}"],
  ["\\Overrightarrow", "\\Overrightarrow{AB}"],
  ["\\overleftrightarrow", "\\overleftrightarrow{AB}"],
  ["\\overleftharpoon", "\\overleftharpoon{AB}"],
  ["\\overrightharpoon", "\\overrightharpoon{AB}"],
  ["\\overlinesegment", "\\overlinesegment{AB}"],
  ["\\underarc", "\\underarc{AB}"],
  ["\\underline", "\\underline{AB}"],
  ["\\underbrace", "\\underbrace{a+b}"],
  ["\\undergroup", "\\undergroup{AB}"],
  ["\\underparen", "\\underparen{AB}"],
  ["\\underleftarrow", "\\underleftarrow{AB}"],
  ["\\underrightarrow", "\\underrightarrow{AB}"],
  ["\\underleftrightarrow", "\\underleftrightarrow{AB}"],
  ["\\underlinesegment", "\\underlinesegment{AB}"],
  ["\\overset", "\\overset{a}{b}"],
  ["\\underset", "\\underset{a}{b}"],
  ["\\overunderset", "\\overunderset{a}{c}{b}"],
  ["\\stackrel", "\\stackrel{a}{b}"],
  ["\\stackbin", "\\stackbin{a}{b}"],
  ["\\cancel", "\\cancel{x}"],
  ["\\bcancel", "\\bcancel{x}"],
  ["\\xcancel", "\\xcancel{x}"],
  ["\\enclose", "\\enclose{circle}{x}"],
  ["\\angl", "\\angl{n}"],
  ["\\phase", "\\phase{z}"],
  ["\\ang", "\\ang{45}"],
  ["\\c", "\\c{c}"],
  ["\\ce", "\\ce{H2O}"],
  ["\\pu", "\\pu{9.8 m/s^2}"],
  ["\\pmod", "\\pmod{n}"],
  ["\\mod", "\\mod{n}"],
  ["\\bra", "\\bra{\\psi}"],
  ["\\ket", "\\ket{\\psi}"],
  ["\\braket", "\\braket{\\phi|\\psi}"],
  ["\\set", "\\set{x}"],
  ["\\Bra", "\\Bra{\\psi}"],
  ["\\Ket", "\\Ket{\\psi}"],
  ["\\Braket", "\\Braket{\\phi|\\psi}"],
  ["\\Set", "\\Set{x}"],
  ["\\boxed", "\\boxed{x}"],
  ["\\bbox", "\\bbox{x}"],
  ["\\error", "\\error{x}"],
  ["\\ensuremath", "\\ensuremath{x}"],
  ["\\mathtip", "\\mathtip{x}{tip}"],
  ["\\texttip", "\\texttip{x}{tip}"],
  ["\\color", "{\\color{red}x}"],
  ["\\textcolor", "\\textcolor{red}{x}"],
  ["\\colorbox", "\\colorbox{yellow}{x}"],
  ["\\fcolorbox", "\\fcolorbox{red}{yellow}{x}"],
  ["\\class", "\\class{preview}{x}"],
  ["\\htmlClass", "\\htmlClass{preview}{x}"],
  ["\\cssId", "\\cssId{preview}{x}"],
  ["\\htmlId", "\\htmlId{preview}{x}"],
  ["\\htmlData", "\\htmlData{preview=value}{x}"],
  ["\\style", "\\style{color:red}{x}"],
  ["\\htmlStyle", "\\htmlStyle{color:red}{x}"],
  ["\\href", "\\href{}{x}"],
  ["\\mbox", "\\mbox{ABC}"],
  ["\\text", "\\text{ABC}"],
  ["\\emph", "\\emph{ABC}"],
  ["\\mathchoice", "\\mathchoice{D}{T}{S}{SS}"],
  ["\\mathop", "\\mathop{f}"],
  ["\\mathbin", "a\\mathbin{+}b"],
  ["\\mathrel", "a\\mathrel{=}b"],
  ["\\mathopen", "\\mathopen{(}x"],
  ["\\mathclose", "x\\mathclose{)}"],
  ["\\mathpunct", "a\\mathpunct{,}b"],
  ["\\mathord", "\\mathord{x}"],
  ["\\mathinner", "\\mathinner{x}"],
  ["\\operatorname", "\\operatorname{op}(x)"],
  ["\\operatorname*", "\\operatorname*{lim}_{n}"],
  ["\\not", "\\not{=}"],
  ["\\smash", "\\smash{\\frac{a}{b}}"],
  ["\\rlap", "\\rlap{A}B"],
  ["\\llap", "A\\llap{B}"],
  ["\\mathrlap", "\\mathrlap{A}B"],
  ["\\mathllap", "A\\mathllap{B}"],
  ["\\raisebox", "\\raisebox{0.4em}{x}"],
  ["\\raise", "\\raise{0.4em}{x}"],
  ["\\lower", "\\lower{0.4em}{x}"],
  ["\\rule", "\\rule{1em}{0.08em}"],
]);

for (const command of [
  "xrightarrow",
  "xleftarrow",
  "xRightarrow",
  "xLeftarrow",
  "xleftharpoonup",
  "xleftharpoondown",
  "xrightharpoonup",
  "xrightharpoondown",
  "xlongequal",
  "xtwoheadleftarrow",
  "xtwoheadrightarrow",
  "xleftrightarrow",
  "xLeftrightarrow",
  "xrightleftharpoons",
  "xleftrightharpoons",
  "xhookleftarrow",
  "xhookrightarrow",
  "xmapsto",
  "xtofrom",
  "xleftrightarrows",
  "xRightleftharpoons",
  "xLeftrightharpoons",
]) {
  previewRegistry.set(`\\${command}`, {
    latex: `\\${command}{AB}`,
    kind: "arguments",
  });
}

registerPreviews("state", [
  ["\\displaystyle", "{\\displaystyle \\frac{a}{b}}"],
  ["\\textstyle", "{\\textstyle \\frac{a}{b}}"],
  ["\\scriptstyle", "{\\scriptstyle \\frac{a}{b}}"],
  ["\\scriptscriptstyle", "{\\scriptscriptstyle \\frac{a}{b}}"],
  ["\\tiny", "{\\tiny ABC}"],
  ["\\scriptsize", "{\\scriptsize ABC}"],
  ["\\footnotesize", "{\\footnotesize ABC}"],
  ["\\small", "{\\small ABC}"],
  ["\\normalsize", "{\\normalsize ABC}"],
  ["\\large", "{\\large ABC}"],
  ["\\Large", "{\\Large ABC}"],
  ["\\LARGE", "{\\LARGE ABC}"],
  ["\\huge", "{\\huge ABC}"],
  ["\\Huge", "{\\Huge ABC}"],
  ["\\bfseries", "{\\bfseries ABC}"],
  ["\\mdseries", "{\\mdseries ABC}"],
  ["\\upshape", "{\\upshape ABC}"],
  ["\\slshape", "{\\slshape ABC}"],
  ["\\scshape", "{\\scshape ABC}"],
  ["\\rmfamily", "{\\rmfamily ABC}"],
  ["\\sffamily", "{\\sffamily ABC}"],
  ["\\ttfamily", "{\\ttfamily ABC}"],
  ["\\em", "{\\em ABC}"],
  ["\\fontseries", "\\text{\\fontseries{b}ABC}"],
  ["\\fontshape", "\\text{\\fontshape{it}ABC}"],
  ["\\fontfamily", "\\text{\\fontfamily{sans-serif}ABC}"],
  ["\\selectfont", "\\text{ABC}"],
  ["\\strut", "A\\strut B"],
  ["\\phantom", "A\\phantom{B}C"],
  ["\\hphantom", "A\\hphantom{B}C"],
  ["\\vphantom", "A\\vphantom{\\frac{a}{b}}B"],
]);

for (const command of compatibilityCommands) {
  previewRegistry.set(command.command, {
    latex: command.previewLatex,
    kind: "arguments",
  });
}

// MathLive's native completion popover renders extensible operators in display
// style. For VisualTeX's custom esint/STIX integral SVGs that makes the glyphs
// roughly 2-3x taller than neighboring candidates. Keep the actual editor
// semantics untouched and render completion previews in text style; esint
// outline correctness is handled separately by the generated glyph registry.
for (const command of EXTENDED_INTEGRAL_COMMANDS) {
  previewRegistry.set(`\\${command}`, {
    latex: `{\\textstyle \\${command}}`,
    kind: "operator",
  });
}

registerPreviews("alias", [
  ["\\bf", "{\\bf ABC}"],
  ["\\it", "{\\it ABC}"],
  ["\\boldsymbol", "\\boldsymbol{\\alpha A}"],
  ["\\bm", "\\bm{\\alpha A}"],
  ["\\bold", "\\bold{\\alpha A}"],
  ["\\Bbb", "\\Bbb{ABC}"],
  ["\\mathbb", "\\mathbb{ABC}"],
  ["\\frak", "\\frak{ABC}"],
  ["\\mathfrak", "\\mathfrak{ABC}"],
]);

registerPreviews("state", [
  ["\\mathbf", "\\mathbf{ABC}"],
  ["\\mathit", "\\mathit{ABC}"],
  ["\\mathnormal", "\\mathnormal{ABC}"],
  ["\\mathbfit", "\\mathbfit{ABC}"],
  ["\\mathrm", "\\mathrm{ABC}"],
  ["\\mathsf", "\\mathsf{ABC}"],
  ["\\mathtt", "\\mathtt{ABC}"],
  ["\\mathcal", "\\mathcal{ABC}"],
  ["\\mathscr", "\\mathscr{gG}"],
  ["\\textbf", "\\textbf{ABC}"],
  ["\\textmd", "\\textmd{ABC}"],
  ["\\textup", "\\textup{ABC}"],
  ["\\textnormal", "\\textnormal{ABC}"],
  ["\\textsl", "\\textsl{ABC}"],
  ["\\textit", "\\textit{ABC}"],
  ["\\textsc", "\\textsc{ABC}"],
  ["\\textrm", "\\textrm{ABC}"],
  ["\\textsf", "\\textsf{ABC}"],
  ["\\texttt", "\\texttt{ABC}"],
]);

for (const command of [
  "enskip",
  "enspace",
  "quad",
  "qquad",
  "space",
  "hspace",
  "hspace*",
  "kern",
  "mkern",
  "mskip",
  "hskip",
  "mspace",
]) {
  const parameter =
    command === "enskip" ||
    command === "enspace" ||
    command === "quad" ||
    command === "qquad" ||
    command === "space"
      ? "{}"
      : "{1em}";
  previewRegistry.set(`\\${command}`, {
    latex: `A\\${command}${parameter}B`,
    kind: "spacing",
  });
}

function delimiterPreview(command: string): NativeSuggestionPreview | null {
  if (command === "\\middle") {
    return {
      latex: "\\left(a\\middle|b\\right)",
      kind: "delimiter",
    };
  }
  if (command === "\\left" || command === "\\right") {
    return {
      latex: "\\left(x\\right)",
      kind: "delimiter",
    };
  }

  const match = command.match(/^\\(big|Big|bigg|Bigg)([lmr])?$/);
  if (!match) return null;
  const position = match[2] ?? "";
  const latex =
    position === "l"
      ? `${command}(`
      : position === "r"
        ? `${command})`
        : position === "m"
          ? `a${command}|b`
          : `${command}|`;
  return { latex, kind: "delimiter" };
}

export function nativeSuggestionPreviewHasVisibleInk(preview: HTMLElement) {
  const rendered =
    preview.querySelector<HTMLElement>(".ML__latex") ?? preview;
  if (rendered.querySelector(".ML__error")) return false;
  const text = (rendered.textContent ?? "")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .trim();
  if (text && text !== ".") return true;
  if (rendered.querySelector("svg, canvas, img, .ML__rule")) return true;

  return [rendered, ...rendered.querySelectorAll<HTMLElement>("*")].some(
    (node) => {
      const style = getComputedStyle(node);
      if (
        style.display === "none" ||
        style.visibility === "hidden" ||
        Number.parseFloat(style.opacity || "1") <= 0
      ) {
        return false;
      }
      if (style.backgroundImage !== "none") return true;
      const hasVisibleBorder = [
        style.borderTopWidth,
        style.borderRightWidth,
        style.borderBottomWidth,
        style.borderLeftWidth,
      ].some((width) => Number.parseFloat(width) > 0);
      if (hasVisibleBorder && style.borderStyle !== "none") return true;
      for (const pseudo of ["::before", "::after"] as const) {
        const pseudoStyle = getComputedStyle(node, pseudo);
        const content = pseudoStyle.content;
        const maskImage = pseudoStyle.maskImage;
        const webkitMaskImage = pseudoStyle.getPropertyValue("-webkit-mask-image");
        if (
          (maskImage && maskImage !== "none") ||
          (webkitMaskImage && webkitMaskImage !== "none")
        ) {
          return true;
        }
        if (content && content !== "none" && content !== "normal" && content !== '\"\"') {
          return true;
        }
      }
      return false;
    },
  );
}

export function resolveNativeSuggestionPreview(
  command: string,
  preview: HTMLElement,
): NativeSuggestionPreview | null {
  const registered = previewRegistry.get(command);
  if (registered) return registered;

  const delimiter = delimiterPreview(command);
  if (delimiter) return delimiter;

  if (nativeSuggestionPreviewHasVisibleInk(preview)) return null;
  return { latex: "\\boxed{?}", kind: "fallback" };
}
