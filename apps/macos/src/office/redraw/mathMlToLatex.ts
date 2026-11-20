const TOKEN_MAP: Record<string, string> = {
  "−": "-",
  "×": "\\times ",
  "÷": "\\div ",
  "±": "\\pm ",
  "∓": "\\mp ",
  "·": "\\cdot ",
  "∗": "\\ast ",
  "∘": "\\circ ",
  "∞": "\\infty ",
  "∂": "\\partial ",
  "∇": "\\nabla ",
  "∑": "\\sum ",
  "∏": "\\prod ",
  "∫": "\\int ",
  "∬": "\\iint ",
  "∭": "\\iiint ",
  "∮": "\\oint ",
  "√": "\\sqrt{}",
  "≈": "\\approx ",
  "≃": "\\simeq ",
  "≅": "\\cong ",
  "≠": "\\ne ",
  "≤": "\\le ",
  "≥": "\\ge ",
  "≪": "\\ll ",
  "≫": "\\gg ",
  "≡": "\\equiv ",
  "∝": "\\propto ",
  "∈": "\\in ",
  "∉": "\\notin ",
  "∋": "\\ni ",
  "⊂": "\\subset ",
  "⊃": "\\supset ",
  "⊆": "\\subseteq ",
  "⊇": "\\supseteq ",
  "∪": "\\cup ",
  "∩": "\\cap ",
  "∧": "\\land ",
  "∨": "\\lor ",
  "¬": "\\neg ",
  "∀": "\\forall ",
  "∃": "\\exists ",
  "∄": "\\nexists ",
  "∅": "\\varnothing ",
  "⊥": "\\perp ",
  "∥": "\\parallel ",
  "←": "\\leftarrow ",
  "→": "\\rightarrow ",
  "↔": "\\leftrightarrow ",
  "⇐": "\\Leftarrow ",
  "⇒": "\\Rightarrow ",
  "⇔": "\\Leftrightarrow ",
  "↦": "\\mapsto ",
  "α": "\\alpha ",
  "β": "\\beta ",
  "γ": "\\gamma ",
  "δ": "\\delta ",
  "ε": "\\epsilon ",
  "ϵ": "\\varepsilon ",
  "ζ": "\\zeta ",
  "η": "\\eta ",
  "θ": "\\theta ",
  "ϑ": "\\vartheta ",
  "ι": "\\iota ",
  "κ": "\\kappa ",
  "λ": "\\lambda ",
  "μ": "\\mu ",
  "ν": "\\nu ",
  "ξ": "\\xi ",
  "π": "\\pi ",
  "ϖ": "\\varpi ",
  "ρ": "\\rho ",
  "ϱ": "\\varrho ",
  "σ": "\\sigma ",
  "ς": "\\varsigma ",
  "τ": "\\tau ",
  "υ": "\\upsilon ",
  "φ": "\\phi ",
  "ϕ": "\\varphi ",
  "χ": "\\chi ",
  "ψ": "\\psi ",
  "ω": "\\omega ",
  "Γ": "\\Gamma ",
  "Δ": "\\Delta ",
  "Θ": "\\Theta ",
  "Λ": "\\Lambda ",
  "Ξ": "\\Xi ",
  "Π": "\\Pi ",
  "Σ": "\\Sigma ",
  "Υ": "\\Upsilon ",
  "Φ": "\\Phi ",
  "Ψ": "\\Psi ",
  "Ω": "\\Omega ",
};

function localName(node: Element) {
  return node.localName || node.tagName.replace(/^.*:/, "");
}

function elementChildren(node: Element) {
  return Array.from(node.children);
}

function escapeText(value: string) {
  return value
    .replace(/\\/g, "\\textbackslash{}")
    .replace(/([{}#%&_])/g, "\\$1");
}

function escapeIdentifier(value: string) {
  return value
    .replace(/\\/g, "\\backslash ")
    .replace(/([{}#%&_])/g, "\\$1");
}

function token(value: string) {
  const trimmed = value.trim();
  if (!trimmed) return "";
  return TOKEN_MAP[trimmed] ?? trimmed.replace(/([{}#%&_])/g, "\\$1");
}

function groupBase(value: string) {
  const trimmed = value.trim();
  if (trimmed.length === 1 || trimmed.startsWith("\\")) return trimmed;
  return `{${trimmed}}`;
}

function delimiter(value: string, left: boolean) {
  const mapped =
    value === "{" ? "\\{" :
    value === "}" ? "\\}" :
    value === "|" ? (left ? "\\lvert" : "\\rvert") :
    value === "‖" ? (left ? "\\lVert" : "\\rVert") :
    value === "〈" || value === "⟨" ? "\\langle" :
    value === "〉" || value === "⟩" ? "\\rangle" :
    value || ".";
  return `${left ? "\\left" : "\\right"}${mapped} `;
}

function children(node: Element) {
  return elementChildren(node).map(convertElement).join("");
}

function convertTable(node: Element) {
  const rows = elementChildren(node)
    .filter((child) => ["mtr", "mlabeledtr"].includes(localName(child)))
    .map((row) => elementChildren(row).map(convertElement).join(" & "));
  return `\\begin{matrix}${rows.join(" \\\\ ")}\\end{matrix}`;
}

function convertFenced(node: Element) {
  const open = node.getAttribute("open") ?? "(";
  const close = node.getAttribute("close") ?? ")";
  const separator = node.getAttribute("separators") ?? ",";
  return `${delimiter(open, true)}${elementChildren(node).map(convertElement).join(separator)}${delimiter(close, false)}`;
}

function convertMultiScripts(node: Element) {
  const values = elementChildren(node);
  if (!values.length) return "";
  let result = groupBase(convertElement(values[0]));
  for (let index = 1; index < values.length && localName(values[index]) !== "mprescripts"; index += 2) {
    const sub = convertElement(values[index]);
    const sup = index + 1 < values.length ? convertElement(values[index + 1]) : "";
    if (sub) result += `_{${sub}}`;
    if (sup) result += `^{${sup}}`;
  }
  return result;
}

function accentCommand(value: string, under: boolean) {
  const commands: Record<string, string> = {
    "¯": "\\bar",
    "̄": "\\bar",
    "̂": "\\hat",
    "˜": "\\tilde",
    "̃": "\\tilde",
    "˙": "\\dot",
    "̇": "\\dot",
    "¨": "\\ddot",
    "̈": "\\ddot",
    "→": "\\vec",
    "⃗": "\\vec",
    "⌢": "\\widehat",
    "⏞": "\\overbrace",
    "⏟": "\\underbrace",
  };
  return commands[value] ?? (under ? "\\underaccent" : "\\overset");
}

function convertAccent(node: Element, under: boolean) {
  const values = elementChildren(node);
  if (!values.length) return "";
  const body = convertElement(values[0]);
  const mark = values.length > 1 ? values[1].textContent?.trim() ?? "" : "";
  const command = accentCommand(mark, under);
  if (command === "\\underaccent" || command === "\\overset") {
    return `${command}{${token(mark)}}{${body}}`;
  }
  return `${command}{${body}}`;
}

function convertElement(node: Element): string {
  const values = elementChildren(node);
  switch (localName(node)) {
    case "math":
    case "mrow":
    case "mstyle":
    case "semantics":
    case "mtd":
      return children(node);
    case "annotation":
    case "annotation-xml":
    case "mspace":
    case "none":
      return "";
    case "mi": {
      const value = node.textContent?.trim() ?? "";
      const variant = node.getAttribute("mathvariant")?.toLowerCase() ?? "";
      if ((variant.includes("normal") || variant.includes("upright")) && /^[A-Za-zÀ-ɏ0-9.,]+$/u.test(value)) {
        return `\\mathrm{${escapeIdentifier(value)}}`;
      }
      return token(value);
    }
    case "mn":
    case "mo":
      return token(node.textContent ?? "");
    case "mtext": {
      const value = node.textContent?.trim() ?? "";
      return value.length === 1 && /^[A-Za-zÀ-ɏ]$/u.test(value)
        ? `\\mathrm{${escapeIdentifier(value)}}`
        : `\\text{${escapeText(node.textContent ?? "")}}`;
    }
    case "mfrac":
      return `\\frac{${values[0] ? convertElement(values[0]) : ""}}{${values[1] ? convertElement(values[1]) : ""}}`;
    case "msqrt":
      return `\\sqrt{${children(node)}}`;
    case "mroot":
      return `\\sqrt[${values[1] ? convertElement(values[1]) : ""}]{${values[0] ? convertElement(values[0]) : ""}}`;
    case "msub":
      return `${groupBase(values[0] ? convertElement(values[0]) : "")}_{${values[1] ? convertElement(values[1]) : ""}}`;
    case "msup":
      return `${groupBase(values[0] ? convertElement(values[0]) : "")}^{${values[1] ? convertElement(values[1]) : ""}}`;
    case "msubsup":
      return `${groupBase(values[0] ? convertElement(values[0]) : "")}_{${values[1] ? convertElement(values[1]) : ""}}^{${values[2] ? convertElement(values[2]) : ""}}`;
    case "munder":
      return `${groupBase(values[0] ? convertElement(values[0]) : "")}_{${values[1] ? convertElement(values[1]) : ""}}`;
    case "mover":
      return convertAccent(node, false);
    case "munderover":
      return `${groupBase(values[0] ? convertElement(values[0]) : "")}_{${values[1] ? convertElement(values[1]) : ""}}^{${values[2] ? convertElement(values[2]) : ""}}`;
    case "mfenced":
      return convertFenced(node);
    case "mtable":
      return convertTable(node);
    case "menclose": {
      const body = children(node);
      const notation = node.getAttribute("notation")?.toLowerCase() ?? "";
      if (notation.includes("box")) return `\\boxed{${body}}`;
      if (notation.includes("radical")) return `\\sqrt{${body}}`;
      return body;
    }
    case "mphantom":
      return `\\phantom{${children(node)}}`;
    case "mmultiscripts":
      return convertMultiScripts(node);
    default:
      return children(node) || token(node.textContent ?? "");
  }
}

export function mathMlToLatex(mathMl: string) {
  if (!mathMl || mathMl.length > 4_000_000 || /<!DOCTYPE|<!ENTITY/i.test(mathMl)) {
    throw new Error("Word returned invalid or excessive MathML.");
  }
  const document = new DOMParser().parseFromString(mathMl, "application/xml");
  const parserError = document.querySelector("parsererror");
  if (parserError) throw new Error("Word MathML could not be parsed.");
  const root = document.documentElement;
  const result = convertElement(root).trim().replace(/ {2,}/g, " ");
  if (!result) throw new Error("Word MathML did not contain a recoverable formula.");
  return result;
}
