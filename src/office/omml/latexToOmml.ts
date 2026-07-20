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

export type OmmlDisplayMode = "inline" | "block";

export interface OmmlArtifacts {
  omml: string;
  ommlBase64: string;
  ommlDocxBase64: string;
}

const MATH_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/math";
const WORD_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);
const texInput = new TeX({
  packages: AllPackages,
  formatError: (_jax: unknown, error: unknown) => {
    throw error instanceof Error ? error : new Error(String(error));
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
  "∫",
  "∬",
  "∭",
  "∮",
  "∯",
  "∰",
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

function normalizeLines(lines: string[]) {
  const normalized = lines
    .map((line) => line.replace(/\r\n?/g, "\n").trim())
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

type RunStyle = "math" | "plain" | "normal";

function ommlRun(value: string, style: RunStyle = "math") {
  const text = sanitizeXmlText(value);
  if (!text) return "";
  const properties =
    style === "plain"
      ? '<m:rPr><m:sty m:val="p"/></m:rPr>'
      : style === "normal"
        ? "<m:rPr><m:nor/></m:rPr>"
        : "";
  const preserve = /^\s|\s$|\s{2,}/.test(text)
    ? ' xml:space="preserve"'
    : "";
  return `<m:r>${properties}<m:t${preserve}>${escapeXmlText(text)}</m:t></m:r>`;
}

function elementChildren(element: Element) {
  return Array.from(element.children);
}

function elementName(element: Element) {
  return element.localName.toLowerCase();
}

function latexToMathMl(latex: string, displayMode: OmmlDisplayMode) {
  const root = mathDocument.convert(latex, {
    display: displayMode === "block",
    end: STATE.COMPILED,
  }) as MmlNode;
  return serializedMmlVisitor.visitTree(root);
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

function tokenRunStyle(element: Element): RunStyle {
  const name = elementName(element);
  if (name === "mtext" || name === "ms") return "normal";
  if (name === "mn" || name === "mo") return "plain";
  const variant = element.getAttribute("mathvariant")?.toLowerCase() ?? "";
  if (
    variant.includes("normal") ||
    variant.includes("upright") ||
    variant.includes("sans-serif") ||
    variant.includes("monospace")
  ) {
    return "plain";
  }
  return "math";
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
  return (
    "<m:acc><m:accPr>" +
    `<m:chr m:val="${escapeXmlAttribute(character)}"/>` +
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
    if (OVER_GROUP_CHARACTERS.has(character)) {
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

function convertTable(element: Element) {
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
      `${index > 0 ? ommlRun(separator, "plain") : ""}${convertElement(child)}`,
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
      return ommlRun(normalizedTokenText(element), tokenRunStyle(element));
    case "mspace":
      return ommlRun(mspaceText(element), "plain");
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
      return ommlRun(normalizedTokenText(element), "math");
    }
  }
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

function minimalDocxBytes(omml: string) {
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
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:m="${MATH_NAMESPACE}">` +
    `<w:body><w:p>${omml}</w:p><w:sectPr/></w:body>` +
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
) {
  const normalized = normalizeLines(lines);
  const converted = normalized.map((line) => {
    const mathElement = parseMathMl(line, displayMode);
    return convertSequence(elementChildren(mathElement));
  });
  const body =
    converted.length === 1
      ? converted[0]
      : `<m:eqArr>${converted
          .map((line) => `<m:e>${line}</m:e>`)
          .join("")}</m:eqArr>`;
  return wrapOmml(body);
}

export function latexLinesToOmmlArtifacts(
  lines: string[],
  displayMode: OmmlDisplayMode,
): OmmlArtifacts {
  const omml = latexLinesToOmml(lines, displayMode);
  return {
    omml,
    ommlBase64: utf8ToBase64Url(omml),
    ommlDocxBase64: bytesToBase64Url(minimalDocxBytes(omml)),
  };
}
