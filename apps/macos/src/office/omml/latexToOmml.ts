import { zipSync, strToU8 } from "fflate";
import { mathjax } from "mathjax-full/js/mathjax.js";
import { TeX } from "mathjax-full/js/input/tex.js";
import { SVG } from "mathjax-full/js/output/svg.js";
import { liteAdaptor } from "mathjax-full/js/adaptors/liteAdaptor.js";
import { RegisterHTMLHandler } from "mathjax-full/js/handlers/html.js";
import { AllPackages } from "mathjax-full/js/input/tex/AllPackages.js";
import { STATE } from "mathjax-full/js/core/MathItem.js";
import { SerializedMmlVisitor } from "mathjax-full/js/core/MmlTree/SerializedMmlVisitor.js";
import type { MmlNode } from "mathjax-full/js/core/MmlTree/MmlNode.js";
import { normalizeMathLiveCanonicalUprightCommands } from "../../editor/normalizeChineseLatex.ts";
import {
  stripVisualTexAlignmentMarkers,
  VISUALTEX_ALIGNMENT_MARKER_LATEX,
} from "../../editor/alignmentMarkers";
import {
  MATHJAX_INTEGRAL_OPERATOR_CHARACTERS,
  normalizeMathJaxUnsupportedNaryCommands,
} from "../../export/mathJaxCompatibility.ts";
import type { LatexCodeFormat } from "../../types/formula";
import {
  assertResolvedPresentationMathMl,
  VISUALTEX_MATHML_MACROS,
} from "../../math/latexCompatibility.ts";
import { errorMessage } from "../../runtime/errorMessage";
import { expandCustomSymbolsForMathMl } from "../../math/customSymbolRendering";
import {
  DEFAULT_FORMULA_CHINESE_FONT,
  DEFAULT_FORMULA_LETTER_FONT,
  formulaChinesePrimaryFontName,
  formulaLetterPrimaryFontName,
  normalizeFormulaChineseFont,
  normalizeFormulaLetterFont,
  type FormulaChineseFont,
  type FormulaLetterFont,
} from "../../editor/formulaFontPreferences";

export type OmmlDisplayMode = "inline" | "block";

export interface OmmlArtifacts {
  omml: string;
  ommlBase64: string;
  ommlDocxBase64: string;
}

export interface OmmlFontPreferences {
  formulaLetterFont?: FormulaLetterFont;
  formulaChineseFont?: FormulaChineseFont;
}

const MATH_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/math";
const WORD_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);
const texInput = new TeX({
  packages: AllPackages,
  macros: VISUALTEX_MATHML_MACROS,
  formatError: (_jax: unknown, error: unknown) => {
    throw new Error(
      errorMessage(error, "MathJax 无法解析该公式。"),
      { cause: error },
    );
  },
});
const mathDocument = mathjax.document("", {
  InputJax: texInput,
  OutputJax: new SVG({
    fontCache: "none",
    internalSpeechTitles: false,
  }),
});
const serializedMmlVisitor = new SerializedMmlVisitor(mathDocument.mmlFactory);

const NARY_OPERATORS = new Set([
  ...MATHJAX_INTEGRAL_OPERATOR_CHARACTERS,
  "∑",
  "∏",
  "∐",
  "⋂",
  "⋃",
  "⨀",
  "⨁",
  "⨂",
  "⨄",
  "⨆",
]);

const HARD_SEQUENCE_BOUNDARIES = new Set([
  "=",
  "≠",
  "<",
  ">",
  "≤",
  "≥",
  "≈",
  "≃",
  "≅",
  "≡",
  "∼",
  "∝",
  "∈",
  "∉",
  "⊂",
  "⊃",
  "⊆",
  "⊇",
  "→",
  "←",
  "⇒",
  "⇐",
  "⇔",
  ",",
  ";",
]);

const OPEN_DELIMITERS = new Set([
  "(",
  "[",
  "{",
  "⟨",
  "⌈",
  "⌊",
  "⟦",
  "|",
  "‖",
  "",
]);
const CLOSE_DELIMITERS = new Set([
  ")",
  "]",
  "}",
  "⟩",
  "⌉",
  "⌋",
  "⟧",
  "|",
  "‖",
  "",
]);

const OVER_BAR_CHARACTERS = new Set(["―", "¯", "‾", "_"]);
const UNDER_BAR_CHARACTERS = new Set(["_", "―", "¯", "‾"]);
const OVER_GROUP_CHARACTERS = new Set(["⏞", "︷", "︵"]);
const UNDER_GROUP_CHARACTERS = new Set(["⏟", "︸", "︶"]);
const ACCENT_CHARACTERS = new Set([
  "^",
  "~",
  "˙",
  "¨",
  "´",
  "`",
  "ˇ",
  "˘",
  "→",
  "←",
  "↔",
  "⃗",
  "̂",
  "̃",
  "̇",
  "̈",
]);
const STRETCHY_ARROW_CHARACTERS = new Set(["→", "←", "↔"]);
const OMML_ACCENT_CHARACTER_OVERRIDES = new Map([
  // A spacing ASCII caret touches the top-left edge of italic letters in Word.
  // The modifier circumflex remains a distinct, centred hat without the wide
  // rightward displacement produced by a combining U+0302 accent.
  ["^", "ˆ"],
]);

function normalizeLines(lines: string[]) {
  const normalized = lines
    .map((line) =>
      normalizeMathJaxUnsupportedNaryCommands(
        normalizeMathLiveCanonicalUprightCommands(
          line.replace(/\r\n?/g, "\n"),
        ),
      ).trim(),
    )
    .filter(Boolean);
  if (normalized.length === 0) {
    throw new Error("Cannot generate Word OMML for an empty formula.");
  }
  return normalized;
}

function sanitizeXmlText(value: string) {
  let output = "";
  for (const character of value) {
    const code = character.codePointAt(0) ?? 0;
    if (
      code === 0x9 ||
      code === 0xa ||
      code === 0xd ||
      (code >= 0x20 && code <= 0xd7ff) ||
      (code >= 0xe000 && code <= 0xfffd) ||
      (code >= 0x10000 && code <= 0x10ffff)
    ) {
      output += character;
    }
  }
  return output;
}

function escapeXmlText(value: string) {
  return sanitizeXmlText(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeXmlAttribute(value: string) {
  return escapeXmlText(value)
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function normalizedTokenText(element: Element) {
  return sanitizeXmlText(element.textContent ?? "")
    .replace(/[\u2061-\u2064\ufeff]/g, "")
    .replaceAll("\u00a0", " ");
}

type OmmlScript =
  | "roman"
  | "script"
  | "fraktur"
  | "double-struck"
  | "sans-serif"
  | "monospace";

type OmmlStyle = "p" | "b" | "i" | "bi";

interface OmmlRunProperties {
  equationArrayAlignment?: boolean;
  normalText?: boolean;
  script?: OmmlScript;
  style?: OmmlStyle;
}

const MATH_VARIANT_RUN_PROPERTIES: Record<string, OmmlRunProperties> = {
  normal: { normalText: true, script: "roman", style: "p" },
  upright: { normalText: true, script: "roman", style: "p" },
  bold: { script: "roman", style: "b" },
  italic: { script: "roman", style: "i" },
  "bold-italic": { script: "roman", style: "bi" },
  "double-struck": { script: "double-struck", style: "p" },
  script: { script: "script", style: "p" },
  "bold-script": { script: "script", style: "b" },
  fraktur: { script: "fraktur", style: "p" },
  "bold-fraktur": { script: "fraktur", style: "b" },
  "sans-serif": { script: "sans-serif", style: "p" },
  "bold-sans-serif": { script: "sans-serif", style: "b" },
  "sans-serif-italic": { script: "sans-serif", style: "i" },
  "sans-serif-bold-italic": { script: "sans-serif", style: "bi" },
  monospace: { script: "monospace", style: "p" },
};

function ommlRun(value: string, properties: OmmlRunProperties = {}) {
  // OMML equation arrays use literal ampersands as non-rendering alignment
  // controls: every odd ampersand is an alignment point and the beginning of
  // each row is the implied spacer. m:aln is a different mechanism and Word
  // does not use it to align rows inside m:eqArr.
  const text = sanitizeXmlText(
    `${properties.equationArrayAlignment ? "&" : ""}${value}`,
  );
  if (!text) return "";
  const propertyBody = [
    properties.normalText ? "<m:nor/>" : "",
    properties.script
      ? `<m:scr m:val="${escapeXmlAttribute(properties.script)}"/>`
      : "",
    properties.style
      ? `<m:sty m:val="${escapeXmlAttribute(properties.style)}"/>`
      : "",
  ].join("");
  const runProperties = propertyBody ? `<m:rPr>${propertyBody}</m:rPr>` : "";
  const preserve = /^\s|\s$|\s{2,}/.test(text)
    ? ' xml:space="preserve"'
    : "";
  return `<m:r>${runProperties}<m:t${preserve}>${escapeXmlText(text)}</m:t></m:r>`;
}

function elementChildren(element: Element) {
  return Array.from(element.children);
}

function elementName(element: Element) {
  return element.localName.toLowerCase();
}

function latexToMathMl(latex: string, displayMode: OmmlDisplayMode) {
  const semanticLatex = expandCustomSymbolsForMathMl(latex);
  const root = mathDocument.convert(semanticLatex, {
    display: displayMode === "block",
    end: STATE.COMPILED,
  }) as MmlNode;
  const mathMl = serializedMmlVisitor.visitTree(root);
  assertResolvedPresentationMathMl(mathMl);
  return mathMl;
}

function parseMathMl(latex: string, displayMode: OmmlDisplayMode) {
  if (typeof DOMParser === "undefined") {
    throw new Error("Word OMML export requires a browser DOM parser.");
  }
  const mathMl = latexToMathMl(latex, displayMode);
  const documentObject = new DOMParser().parseFromString(
    mathMl,
    "application/xml",
  );
  const parseError = documentObject.querySelector("parsererror");
  if (parseError) {
    throw new Error(
      `MathJax produced invalid MathML: ${parseError.textContent ?? "parse error"}`,
    );
  }
  if (documentObject.documentElement.localName !== "math") {
    throw new Error("MathJax did not produce a MathML math element.");
  }
  return documentObject.documentElement;
}

const RELATION_ALIGNMENT_TOKENS = new Set([
  "=",
  "≠",
  "<",
  ">",
  "≤",
  "≥",
  "≈",
  "≃",
  "≅",
  "≡",
  "∼",
  "∝",
  "∈",
  "∉",
  "⊂",
  "⊃",
  "⊆",
  "⊇",
  "→",
  "←",
  "⇒",
  "⇐",
  "⇔",
]);

const ALIGNMENT_TRANSPARENT_ELEMENTS = new Set([
  "math",
  "mrow",
  "mstyle",
  "mpadded",
  "maction",
  "semantics",
  "mtd",
]);

const ALIGNMENT_ATTRIBUTE = "data-visualtex-omml-alignment";

function findTopLevelRelationElement(element: Element): Element | null {
  const name = elementName(element);
  if (
    name === "mo" &&
    RELATION_ALIGNMENT_TOKENS.has(normalizedTokenText(element).trim())
  ) {
    return element;
  }
  if (!ALIGNMENT_TRANSPARENT_ELEMENTS.has(name)) return null;
  for (const child of elementChildren(element)) {
    if (["annotation", "annotation-xml"].includes(elementName(child))) continue;
    const relation = findTopLevelRelationElement(child);
    if (relation) return relation;
  }
  return null;
}

function markTopLevelRelationAlignment(mathElement: Element) {
  findTopLevelRelationElement(mathElement)?.setAttribute(
    ALIGNMENT_ATTRIBUTE,
    "true",
  );
}

function relationAlignedCodeFormat(codeFormat: string) {
  return [
    "align",
    "align-star",
    "aligned",
    "equation-split",
    "equation-star-split",
  ].includes(codeFormat);
}

function effectiveMathVariant(element: Element) {
  let current: Element | null = element;
  while (current) {
    const variant = current.getAttribute("mathvariant")?.trim();
    if (variant) {
      return variant.toLowerCase().replace(/[\s_]+/g, "-");
    }
    current = current.parentElement;
  }
  return "";
}

function tokenRunProperties(element: Element): OmmlRunProperties {
  const name = elementName(element);
  const equationArrayAlignment =
    element.getAttribute(ALIGNMENT_ATTRIBUTE) === "true";
  const variant = effectiveMathVariant(element);
  const explicitProperties = variant
    ? MATH_VARIANT_RUN_PROPERTIES[variant]
    : undefined;

  if (name === "mtext" || name === "ms") {
    const properties = explicitProperties
      ? { ...explicitProperties, normalText: true }
      : { normalText: true };
    return equationArrayAlignment
      ? { ...properties, equationArrayAlignment }
      : properties;
  }
  if (explicitProperties) {
    return equationArrayAlignment
      ? { ...explicitProperties, equationArrayAlignment }
      : explicitProperties;
  }

  if (name === "mn" || name === "mo") {
    return {
      script: "roman",
      style: "p",
      ...(equationArrayAlignment ? { equationArrayAlignment } : {}),
    };
  }
  if (name === "mi") {
    const tokenLength = Array.from(normalizedTokenText(element).trim()).length;
    const properties: OmmlRunProperties = tokenLength > 1
      ? { script: "roman", style: "p" }
      : {};
    return equationArrayAlignment
      ? { ...properties, equationArrayAlignment }
      : properties;
  }
  return equationArrayAlignment ? { equationArrayAlignment } : {};
}

function mspaceText(element: Element) {
  const width = element.getAttribute("width")?.trim().toLowerCase() ?? "";
  if (!width || width.startsWith("0") || width.startsWith("-")) return "";
  const numeric = Number.parseFloat(width);
  if (!Number.isFinite(numeric)) return " ";
  if (width.endsWith("em")) {
    if (numeric <= 0.2) return "\u2009";
    if (numeric <= 0.35) return "\u2005";
    if (numeric <= 0.6) return "\u2004";
    return "\u2003";
  }
  return " ";
}

function delimiterFromElement(
  element: Element | undefined,
  kind: "open" | "close",
): string | null {
  if (!element) return null;
  const name = elementName(element);
  if (name === "mo") {
    const value = normalizedTokenText(element).trim();
    const texClass =
      element.getAttribute("data-mjx-texclass")?.toUpperCase() ?? "";
    const allowed = kind === "open" ? OPEN_DELIMITERS : CLOSE_DELIMITERS;
    if (
      allowed.has(value) &&
      (texClass === "" ||
        texClass === (kind === "open" ? "OPEN" : "CLOSE") ||
        value === "|" ||
        value === "‖")
    ) {
      return value;
    }
    return null;
  }
  const children = elementChildren(element);
  if (children.length !== 1) return null;
  const texClass =
    element.getAttribute("data-mjx-texclass")?.toUpperCase() ?? "";
  if (
    texClass &&
    texClass !== (kind === "open" ? "OPEN" : "CLOSE")
  ) {
    return null;
  }
  return delimiterFromElement(children[0], kind);
}

function ommlDelimiter(begin: string, end: string, body: string) {
  return (
    "<m:d><m:dPr>" +
    `<m:begChr m:val="${escapeXmlAttribute(begin)}"/>` +
    `<m:endChr m:val="${escapeXmlAttribute(end)}"/>` +
    "</m:dPr>" +
    `<m:e>${body}</m:e>` +
    "</m:d>"
  );
}

function isHardSequenceBoundary(element: Element) {
  return (
    elementName(element) === "mo" &&
    HARD_SEQUENCE_BOUNDARIES.has(normalizedTokenText(element).trim())
  );
}

interface NaryParts {
  character: string;
  subscript?: Element;
  superscript?: Element;
  limitLocation: "subSup" | "undOvr";
}

function naryParts(element: Element): NaryParts | null {
  const name = elementName(element);
  if (name === "mo") {
    const character = normalizedTokenText(element).trim();
    return NARY_OPERATORS.has(character)
      ? { character, limitLocation: "subSup" }
      : null;
  }
  const children = elementChildren(element);
  const base = children[0];
  if (!base || elementName(base) !== "mo") return null;
  const character = normalizedTokenText(base).trim();
  if (!NARY_OPERATORS.has(character)) return null;

  switch (name) {
    case "msub":
      return {
        character,
        subscript: children[1],
        limitLocation: "subSup",
      };
    case "msup":
      return {
        character,
        superscript: children[1],
        limitLocation: "subSup",
      };
    case "msubsup":
      return {
        character,
        subscript: children[1],
        superscript: children[2],
        limitLocation: "subSup",
      };
    case "munder":
      return {
        character,
        subscript: children[1],
        limitLocation: "undOvr",
      };
    case "mover":
      return {
        character,
        superscript: children[1],
        limitLocation: "undOvr",
      };
    case "munderover":
      return {
        character,
        subscript: children[1],
        superscript: children[2],
        limitLocation: "undOvr",
      };
    default:
      return null;
  }
}

function convertNary(element: Element, bodyElements: Element[]) {
  const parts = naryParts(element);
  if (!parts) return "";
  const subscript = parts.subscript ? convertElement(parts.subscript) : "";
  const superscript = parts.superscript
    ? convertElement(parts.superscript)
    : "";
  const body = convertSequence(bodyElements);
  return (
    "<m:nary><m:naryPr>" +
    `<m:chr m:val="${escapeXmlAttribute(parts.character)}"/>` +
    `<m:limLoc m:val="${parts.limitLocation}"/>` +
    `<m:subHide m:val="${parts.subscript ? "0" : "1"}"/>` +
    `<m:supHide m:val="${parts.superscript ? "0" : "1"}"/>` +
    "</m:naryPr>" +
    `<m:sub>${subscript}</m:sub>` +
    `<m:sup>${superscript}</m:sup>` +
    `<m:e>${body}</m:e>` +
    "</m:nary>"
  );
}

function convertSequence(elements: Element[]) {
  let output = "";
  for (let index = 0; index < elements.length; index += 1) {
    const element = elements[index];
    if (naryParts(element)) {
      let bodyEnd = index + 1;
      while (
        bodyEnd < elements.length &&
        !isHardSequenceBoundary(elements[bodyEnd])
      ) {
        bodyEnd += 1;
      }
      output += convertNary(element, elements.slice(index + 1, bodyEnd));
      index = bodyEnd - 1;
      continue;
    }
    output += convertElement(element);
  }
  return output;
}

function convertMrow(element: Element) {
  const children = elementChildren(element);
  if (children.length >= 1) {
    const begin = delimiterFromElement(children[0], "open");
    const end = delimiterFromElement(children.at(-1), "close");
    if (begin !== null && end !== null && children.length >= 2) {
      return ommlDelimiter(begin, end, convertSequence(children.slice(1, -1)));
    }
    if (begin !== null && children.length >= 2) {
      return ommlDelimiter(begin, "", convertSequence(children.slice(1)));
    }
  }
  return convertSequence(children);
}

function convertFraction(element: Element) {
  const children = elementChildren(element);
  const numerator = children[0] ? convertElement(children[0]) : "";
  const denominator = children[1] ? convertElement(children[1]) : "";
  const thickness = element.getAttribute("linethickness")?.trim() ?? "";
  const bevelled = element.getAttribute("bevelled") === "true";
  const type = bevelled
    ? "skw"
    : thickness === "0" || thickness === "0px" || thickness === "0em"
      ? "noBar"
      : "bar";
  return (
    `<m:f><m:fPr><m:type m:val="${type}"/></m:fPr>` +
    `<m:num>${numerator}</m:num>` +
    `<m:den>${denominator}</m:den>` +
    "</m:f>"
  );
}

function convertRadical(element: Element) {
  const children = elementChildren(element);
  if (elementName(element) === "msqrt") {
    return (
      '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>' +
      "<m:deg></m:deg>" +
      `<m:e>${convertSequence(children)}</m:e>` +
      "</m:rad>"
    );
  }
  const radicand = children[0] ? convertElement(children[0]) : "";
  const degree = children[1] ? convertElement(children[1]) : "";
  return (
    '<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr>' +
    `<m:deg>${degree}</m:deg>` +
    `<m:e>${radicand}</m:e>` +
    "</m:rad>"
  );
}

function convertScript(element: Element) {
  if (naryParts(element)) return convertNary(element, []);
  const children = elementChildren(element);
  const base = children[0] ? convertElement(children[0]) : "";
  if (elementName(element) === "msub") {
    return `<m:sSub><m:e>${base}</m:e><m:sub>${children[1] ? convertElement(children[1]) : ""}</m:sub></m:sSub>`;
  }
  if (elementName(element) === "msup") {
    return `<m:sSup><m:e>${base}</m:e><m:sup>${children[1] ? convertElement(children[1]) : ""}</m:sup></m:sSup>`;
  }
  return (
    `<m:sSubSup><m:e>${base}</m:e>` +
    `<m:sub>${children[1] ? convertElement(children[1]) : ""}</m:sub>` +
    `<m:sup>${children[2] ? convertElement(children[2]) : ""}</m:sup>` +
    "</m:sSubSup>"
  );
}

function operatorLooksLikeLimit(element: Element) {
  if (elementName(element) !== "mo") return false;
  const text = normalizedTokenText(element).trim().toLowerCase();
  return (
    element.getAttribute("movablelimits") === "true" ||
    element.getAttribute("data-mjx-texclass") === "OP" ||
    ["lim", "max", "min", "sup", "inf", "det", "gcd", "Pr"].includes(text)
  );
}

function accentCharacter(element: Element | undefined) {
  if (!element) return "";
  const text = normalizedTokenText(element).trim();
  return Array.from(text)[0] ?? "";
}

function convertAccent(base: string, character: string) {
  const ommlCharacter =
    OMML_ACCENT_CHARACTER_OVERRIDES.get(character) ?? character;
  return (
    "<m:acc><m:accPr>" +
    `<m:chr m:val="${escapeXmlAttribute(ommlCharacter)}"/>` +
    "</m:accPr>" +
    `<m:e>${base}</m:e>` +
    "</m:acc>"
  );
}

function convertBar(base: string, position: "top" | "bot") {
  return (
    `<m:bar><m:barPr><m:pos m:val="${position}"/></m:barPr>` +
    `<m:e>${base}</m:e></m:bar>`
  );
}

function convertGroupCharacter(
  base: string,
  character: string,
  position: "top" | "bot",
) {
  return (
    "<m:groupChr><m:groupChrPr>" +
    `<m:chr m:val="${escapeXmlAttribute(character)}"/>` +
    `<m:pos m:val="${position}"/>` +
    `<m:vertJc m:val="${position === "top" ? "bot" : "top"}"/>` +
    "</m:groupChrPr>" +
    `<m:e>${base}</m:e>` +
    "</m:groupChr>"
  );
}

function convertUnderOver(element: Element) {
  if (naryParts(element)) return convertNary(element, []);
  const children = elementChildren(element);
  const baseElement = children[0];
  const base = baseElement ? convertElement(baseElement) : "";
  const name = elementName(element);

  if (name === "mover") {
    const upperElement = children[1];
    const upper = upperElement ? convertElement(upperElement) : "";
    const character = accentCharacter(upperElement);
    if (OVER_BAR_CHARACTERS.has(character)) return convertBar(base, "top");
    if (
      OVER_GROUP_CHARACTERS.has(character) ||
      STRETCHY_ARROW_CHARACTERS.has(character)
    ) {
      // Word's m:acc arrow touches or overlaps the base at normal document
      // sizes. A top group character keeps both short \\vec accents and long
      // \\overrightarrow arrows separate, stretchable and centred.
      return convertGroupCharacter(base, character, "top");
    }
    if (
      ACCENT_CHARACTERS.has(character) ||
      upperElement?.getAttribute("accent") === "true" ||
      element.getAttribute("accent") === "true"
    ) {
      return convertAccent(base, character || "^");
    }
    return `<m:limUpp><m:e>${base}</m:e><m:lim>${upper}</m:lim></m:limUpp>`;
  }

  if (name === "munder") {
    const lowerElement = children[1];
    const lower = lowerElement ? convertElement(lowerElement) : "";
    const character = accentCharacter(lowerElement);
    if (UNDER_BAR_CHARACTERS.has(character)) return convertBar(base, "bot");
    if (UNDER_GROUP_CHARACTERS.has(character)) {
      return convertGroupCharacter(base, character, "bot");
    }
    return `<m:limLow><m:e>${base}</m:e><m:lim>${lower}</m:lim></m:limLow>`;
  }

  const lower = children[1] ? convertElement(children[1]) : "";
  const upper = children[2] ? convertElement(children[2]) : "";
  const upperWrapper = `<m:limUpp><m:e>${base}</m:e><m:lim>${upper}</m:lim></m:limUpp>`;
  if (baseElement && operatorLooksLikeLimit(baseElement)) {
    return `<m:limLow><m:e>${upperWrapper}</m:e><m:lim>${lower}</m:lim></m:limLow>`;
  }
  return `<m:limLow><m:e>${upperWrapper}</m:e><m:lim>${lower}</m:lim></m:limLow>`;
}

function isMathJaxAlignmentTable(element: Element) {
  const columns = (element.getAttribute("columnalign") ?? "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  if (columns.length < 2 || columns.length % 2 !== 0) return false;
  if (!columns.every((value, index) => value === (index % 2 === 0 ? "right" : "left"))) {
    return false;
  }
  const spacing = (element.getAttribute("columnspacing") ?? "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  return spacing.length > 0 && spacing.every((value) => /^0(?:\.0+)?(?:em|ex|px|pt)?$/i.test(value));
}

function convertAlignmentTable(element: Element) {
  const rows = elementChildren(element).filter((child) =>
    ["mtr", "mlabeledtr"].includes(elementName(child)),
  );
  const rowXml = rows
    .map((row) => {
      const cells = elementChildren(row).filter(
        (child) => elementName(child) === "mtd",
      );
      cells.forEach((cell, index) => {
        if (index % 2 === 1) markTopLevelRelationAlignment(cell);
      });
      const body = cells
        .map((cell) => convertSequence(elementChildren(cell)))
        .join("");
      return `<m:e>${body}</m:e>`;
    })
    .join("");
  return `<m:eqArr><m:eqArrPr><m:baseJc m:val="center"/></m:eqArrPr>${rowXml}</m:eqArr>`;
}

function convertTable(element: Element) {
  if (isMathJaxAlignmentTable(element)) return convertAlignmentTable(element);
  const rows = elementChildren(element).filter((child) =>
    ["mtr", "mlabeledtr"].includes(elementName(child)),
  );
  const rowXml = rows
    .map((row) => {
      const cells = elementChildren(row).filter(
        (child) => elementName(child) === "mtd",
      );
      return `<m:mr>${cells
        .map((cell) => `<m:e>${convertSequence(elementChildren(cell))}</m:e>`)
        .join("")}</m:mr>`;
    })
    .join("");
  return `<m:m>${rowXml}</m:m>`;
}

function convertFenced(element: Element) {
  const begin = element.getAttribute("open") ?? "(";
  const end = element.getAttribute("close") ?? ")";
  const separators = element.getAttribute("separators") ?? ",";
  const children = elementChildren(element);
  const separator = Array.from(separators)[0] ?? ",";
  const body = children
    .map((child, index) =>
      `${index > 0 ? ommlRun(separator, { script: "roman", style: "p" }) : ""}${convertElement(child)}`,
    )
    .join("");
  return ommlDelimiter(begin, end, body);
}

function convertEnclose(element: Element) {
  const body = convertSequence(elementChildren(element));
  const notation = element.getAttribute("notation")?.toLowerCase() ?? "";
  if (
    !notation ||
    notation.includes("box") ||
    notation.includes("circle") ||
    notation.includes("roundedbox")
  ) {
    return `<m:borderBox><m:e>${body}</m:e></m:borderBox>`;
  }
  return body;
}

function convertSemantics(element: Element) {
  const content = elementChildren(element).find(
    (child) => !["annotation", "annotation-xml"].includes(elementName(child)),
  );
  return content ? convertElement(content) : "";
}

function convertElement(element: Element): string {
  const name = elementName(element);
  switch (name) {
    case "math":
      return convertSequence(elementChildren(element));
    case "mrow":
      return convertMrow(element);
    case "mi":
    case "mn":
    case "mo":
    case "mtext":
    case "ms":
      return ommlRun(normalizedTokenText(element), tokenRunProperties(element));
    case "mspace":
      return ommlRun(mspaceText(element), { script: "roman", style: "p" });
    case "mfrac":
      return convertFraction(element);
    case "msqrt":
    case "mroot":
      return convertRadical(element);
    case "msub":
    case "msup":
    case "msubsup":
      return convertScript(element);
    case "munder":
    case "mover":
    case "munderover":
      return convertUnderOver(element);
    case "mtable":
      return convertTable(element);
    case "mtr":
    case "mlabeledtr":
      return `<m:mr>${elementChildren(element)
        .filter((child) => elementName(child) === "mtd")
        .map((cell) => `<m:e>${convertSequence(elementChildren(cell))}</m:e>`)
        .join("")}</m:mr>`;
    case "mtd":
      return convertSequence(elementChildren(element));
    case "mfenced":
      return convertFenced(element);
    case "menclose":
      return convertEnclose(element);
    case "mphantom":
      return `<m:phant><m:e>${convertSequence(elementChildren(element))}</m:e></m:phant>`;
    case "semantics":
      return convertSemantics(element);
    case "annotation":
    case "annotation-xml":
    case "maligngroup":
    case "malignmark":
    case "none":
      return "";
    case "mstyle":
    case "mpadded":
    case "maction":
      return convertSequence(elementChildren(element));
    case "merror":
      throw new Error(
        `MathJax could not convert this LaTeX formula: ${normalizedTokenText(element)}`,
      );
    default: {
      const children = elementChildren(element);
      if (children.length > 0) return convertSequence(children);
      return ommlRun(normalizedTokenText(element));
    }
  }
}

function applyOmmlFontPreferences(
  body: string,
  preferences: OmmlFontPreferences = {},
) {
  const letterFont = normalizeFormulaLetterFont(
    preferences.formulaLetterFont ?? DEFAULT_FORMULA_LETTER_FONT,
  );
  const chineseFont = normalizeFormulaChineseFont(
    preferences.formulaChineseFont ?? DEFAULT_FORMULA_CHINESE_FONT,
  );
  const letterFontName = formulaLetterPrimaryFontName(letterFont);
  const chineseFontName = formulaChinesePrimaryFontName(chineseFont);

  return body.replace(/<m:r>([\s\S]*?)<\/m:r>/g, (whole, inner: string) => {
    const text = inner.match(/<m:t(?:\s[^>]*)?>([\s\S]*?)<\/m:t>/)?.[1] ?? "";
    if (!text) return whole;

    const hasLatinOrGreek = /[A-Za-z\u0370-\u03ff\u1f00-\u1fff]/u.test(text);
    const hasChinese = /[\u3400-\u9fff\uf900-\ufaff]/u.test(text);
    if (!hasLatinOrGreek && !hasChinese) return whole;

    const script = inner.match(/<m:scr\s+m:val="([^"]+)"\/>/)?.[1] ?? "";
    const isNormalText = inner.includes("<m:nor/>");
    const isTextRun = isNormalText && !script;
    const isSpecialMathAlphabet = Boolean(script && script !== "roman");

    const fontAttributes: string[] = [];
    if (hasLatinOrGreek) {
      const latinFont = isTextRun
        ? chineseFontName
        : letterFont !== DEFAULT_FORMULA_LETTER_FONT && !isSpecialMathAlphabet
          ? letterFontName
          : "";
      if (latinFont) {
        const escaped = escapeXmlAttribute(latinFont);
        fontAttributes.push(
          `w:ascii="${escaped}"`,
          `w:hAnsi="${escaped}"`,
          `w:cs="${escaped}"`,
        );
      }
    }
    if (hasChinese) {
      fontAttributes.push(
        `w:eastAsia="${escapeXmlAttribute(chineseFontName)}"`,
      );
    }
    if (!fontAttributes.length) return whole;

    const wordRunProperties = `<w:rPr><w:rFonts ${fontAttributes.join(" ")}/></w:rPr>`;
    const nextInner = inner.replace(
      /(<m:t(?:\s[^>]*)?>)/,
      `${wordRunProperties}$1`,
    );
    return `<m:r>${nextInner}</m:r>`;
  });
}

function wrapOmml(body: string) {
  return (
    `<m:oMath xmlns:m="${MATH_NAMESPACE}" xmlns:w="${WORD_NAMESPACE}">` +
    body +
    "</m:oMath>"
  );
}

function bytesToBase64Url(bytes: Uint8Array) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary)
    .replaceAll("+", "-")
    .replaceAll("/", "_")
    .replace(/=+$/g, "");
}

function utf8ToBase64Url(value: string) {
  return bytesToBase64Url(new TextEncoder().encode(value));
}

function minimalDocxBytes(omml: string, displayMode: OmmlDisplayMode) {
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    "</Types>";
  const rootRelationships =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    "</Relationships>";
  const mathBody =
    displayMode === "block"
      ? `<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>${omml}</m:oMathPara>`
      : omml;
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:m="${MATH_NAMESPACE}">` +
    `<w:body><w:p>${mathBody}</w:p><w:sectPr/></w:body>` +
    "</w:document>";

  return zipSync(
    {
      "[Content_Types].xml": strToU8(contentTypes),
      "_rels/.rels": strToU8(rootRelationships),
      "word/document.xml": strToU8(documentXml),
    },
    { level: 6 },
  );
}

export function latexLinesToOmml(
  lines: string[],
  displayMode: OmmlDisplayMode,
  codeFormat: LatexCodeFormat | string = "raw",
  fontPreferences: OmmlFontPreferences = {},
) {
  const normalized = normalizeLines(lines);
  const explicitAlignment = relationAlignedCodeFormat(codeFormat);
  const converted = normalized.map((line) => {
    if (!explicitAlignment) {
      const mathElement = parseMathMl(
        stripVisualTexAlignmentMarkers(line),
        displayMode,
      );
      return convertSequence(elementChildren(mathElement));
    }

    return line
      .split(VISUALTEX_ALIGNMENT_MARKER_LATEX)
      .map((segment, index) => {
        const convertedSegment = segment.trim()
          ? convertSequence(elementChildren(parseMathMl(segment, displayMode)))
          : "";
        return index === 0
          ? convertedSegment
          : ommlRun("", { equationArrayAlignment: true }) + convertedSegment;
      })
      .join("");
  });
  const useEquationArray = converted.length > 1 || explicitAlignment;
  const body =
    !useEquationArray
      ? converted[0]
      : `<m:eqArr><m:eqArrPr><m:baseJc m:val="center"/></m:eqArrPr>${converted
          .map((line) => `<m:e>${line}</m:e>`)
          .join("")}</m:eqArr>`;
  return wrapOmml(applyOmmlFontPreferences(body, fontPreferences));
}

export function latexLinesToOmmlArtifacts(
  lines: string[],
  displayMode: OmmlDisplayMode,
  codeFormat: LatexCodeFormat | string = "raw",
  fontPreferences: OmmlFontPreferences = {},
): OmmlArtifacts {
  const omml = latexLinesToOmml(
    lines,
    displayMode,
    codeFormat,
    fontPreferences,
  );
  return {
    omml,
    ommlBase64: utf8ToBase64Url(omml),
    ommlDocxBase64: bytesToBase64Url(minimalDocxBytes(omml, displayMode)),
  };
}
