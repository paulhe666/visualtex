/**
 * Math Characters output: a formula as Unicode text for fields that show
 * plain text (chats, search boxes, file names). Structure-aware: it reads
 * the MathML that VisualTeX's export path makes from the formula (macros
 * expanded, operators classified), so precedence is explicit rather than
 * guessed from the source.
 *
 *   x^2                 -> x²
 *   \frac{a+b}{c+d}     -> (a + b)/(c + d)
 *   x^{a+b}             -> x^(a + b)
 *   \int_a^b f(x)\,dx   -> ∫ₐᵇ f(x) dx
 *   \sqrt{x+1}          -> √(x + 1)
 *   \mathbb{R}          -> ℝ
 *
 * What plain text cannot hold (a fraction, a matrix, a limit under an
 * operator) is written linearly and listed in `notes`.
 */

export interface UnicodeMath {
  text: string;
  notes: string[];
}

type Node = {
  nodeType: number;
  nodeName: string;
  localName?: string | null;
  childNodes: ArrayLike<Node>;
  textContent: string | null;
  getAttribute?: (name: string) => string | null;
};

const ELEMENT = 1;

const SUPERSCRIPTS: Record<string, string> = {
  "0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴", "5": "⁵", "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹",
  "+": "⁺", "-": "⁻", "−": "⁻", "=": "⁼", "(": "⁽", ")": "⁾",
  a: "ᵃ", b: "ᵇ", c: "ᶜ", d: "ᵈ", e: "ᵉ", f: "ᶠ", g: "ᵍ", h: "ʰ", i: "ⁱ", j: "ʲ", k: "ᵏ", l: "ˡ", m: "ᵐ",
  n: "ⁿ", o: "ᵒ", p: "ᵖ", r: "ʳ", s: "ˢ", t: "ᵗ", u: "ᵘ", v: "ᵛ", w: "ʷ", x: "ˣ", y: "ʸ", z: "ᶻ",
  A: "ᴬ", B: "ᴮ", D: "ᴰ", E: "ᴱ", G: "ᴳ", H: "ᴴ", I: "ᴵ", J: "ᴶ", K: "ᴷ", L: "ᴸ", M: "ᴹ", N: "ᴺ",
  O: "ᴼ", P: "ᴾ", R: "ᴿ", T: "ᵀ", U: "ᵁ", V: "ⱽ", W: "ᵂ",
  "α": "ᵅ", "β": "ᵝ", "γ": "ᵞ", "δ": "ᵟ", "θ": "ᶿ", "φ": "ᵠ", "χ": "ᵡ",
  "′": "′", "″": "″", "∗": "*", "*": "*", "∘": "°", " ": "",
};

const SUBSCRIPTS: Record<string, string> = {
  "0": "₀", "1": "₁", "2": "₂", "3": "₃", "4": "₄", "5": "₅", "6": "₆", "7": "₇", "8": "₈", "9": "₉",
  "+": "₊", "-": "₋", "−": "₋", "=": "₌", "(": "₍", ")": "₎",
  a: "ₐ", e: "ₑ", h: "ₕ", i: "ᵢ", j: "ⱼ", k: "ₖ", l: "ₗ", m: "ₘ", n: "ₙ", o: "ₒ", p: "ₚ", r: "ᵣ",
  s: "ₛ", t: "ₜ", u: "ᵤ", v: "ᵥ", x: "ₓ",
  "β": "ᵦ", "γ": "ᵧ", "ρ": "ᵨ", "φ": "ᵩ", "χ": "ᵪ", " ": "",
};

/** Mathematical alphanumeric styles (\mathbf, \mathbb, ...) by MathML mathvariant. */
const ALPHABETS: Record<
  string,
  { upper?: number; lower?: number; digit?: number; greekUpper?: number; greekLower?: number; exceptions?: Record<string, string> }
> = {
  bold: { upper: 0x1d400, lower: 0x1d41a, digit: 0x1d7ce, greekUpper: 0x1d6a8, greekLower: 0x1d6c2 },
  "bold-italic": { upper: 0x1d468, lower: 0x1d482, greekUpper: 0x1d71c, greekLower: 0x1d736 },
  "double-struck": {
    upper: 0x1d538, lower: 0x1d552, digit: 0x1d7d8,
    exceptions: { C: "ℂ", H: "ℍ", N: "ℕ", P: "ℙ", Q: "ℚ", R: "ℝ", Z: "ℤ" },
  },
  script: {
    upper: 0x1d49c, lower: 0x1d4b6,
    exceptions: { B: "ℬ", E: "ℰ", F: "ℱ", H: "ℋ", I: "ℐ", L: "ℒ", M: "ℳ", R: "ℛ", e: "ℯ", g: "ℊ", o: "ℴ" },
  },
  "bold-script": { upper: 0x1d4d0, lower: 0x1d4ea },
  fraktur: {
    upper: 0x1d504, lower: 0x1d51e,
    exceptions: { C: "ℭ", H: "ℌ", I: "ℑ", R: "ℜ", Z: "ℨ" },
  },
  "bold-fraktur": { upper: 0x1d56c, lower: 0x1d586 },
  "sans-serif": { upper: 0x1d5a0, lower: 0x1d5ba, digit: 0x1d7e2 },
  "bold-sans-serif": { upper: 0x1d5d4, lower: 0x1d5ee, digit: 0x1d7ec, greekUpper: 0x1d756, greekLower: 0x1d770 },
  "sans-serif-bold-italic": { upper: 0x1d63c, lower: 0x1d656, greekUpper: 0x1d790, greekLower: 0x1d7aa },
  "sans-serif-italic": { upper: 0x1d608, lower: 0x1d622 },
  monospace: { upper: 0x1d670, lower: 0x1d68a, digit: 0x1d7f6 },
};

const BINARY = new Set("+−-±∓×÷·∘∙⋅∪∩∧∨⊕⊖⊗⊘⊙∖⋆⋄△▽⊎⊓⊔".split(""));
const RELATIONS = new Set(
  "=≠<>≤≥≦≧≪≫≈≅≡≢∼≃≍≐∝∈∉∋∌⊂⊃⊆⊇⊊⊋⊄⊅⊏⊐⊑⊒⊢⊣⊨⊥∥∦→←↔⇒⇐⇔⟶⟵⟷⟹⟸⟺↦⟼↑↓↗↘≺≻⪯⪰≔≕∣∤:".split(""),
);
const OPENING = new Set("([{⟨⌊⌈|‖".split(""));
const LARGE_OPERATORS = /^[∫∬∭∮∯∰∱∲∳⨋⨌⨍⨎⨏⨐⨑⨒⨓⨔⨕⨖⨗⨘⨙⨚⨛⨜∑∏∐⋃⋂⋁⋀⨁⨂⨀⨄⨆⨊]/;
const INVISIBLE = new Set(["\u2061", "\u2062", "\u2063", "\u2064"]);
/** Combining marks for what stands under a character (\\underline,
 *  \\underrightarrow). */
const UNDER_ACCENTS: Record<string, string> = {
  "\u2015": "\u0332", "\u203e": "\u0332", "_": "\u0332", "\u00af": "\u0332",
  "\u2192": "\u20ef", "\u2190": "\u20ee",
  "~": "\u0330", "\u02dc": "\u0330", ".": "\u0323", "\u02d9": "\u0323",
};

/** Combining marks for accents over a single character. */
const ACCENTS: Record<string, string> = {
  "\u2015": "\u0305", // \overline in MathJax's MathML
  "^": "\u0302", "\u02c6": "\u0302", "\u0302": "\u0302",
  "~": "\u0303", "\u02dc": "\u0303", "\u0303": "\u0303",
  "\u00af": "\u0304", "\u203e": "\u0305", "_": "\u0332", "\u0304": "\u0304", "\u0305": "\u0305",
  "\u2192": "\u20d7", "\u20d7": "\u20d7", "\u2190": "\u20d6", "\u2194": "\u20e1",
  "\u20db": "\u20db", "\u20dc": "\u20dc",
  "\u02d9": "\u0307", ".": "\u0307", "\u0307": "\u0307",
  "\u00a8": "\u0308", "\u0308": "\u0308",
  "\u02c7": "\u030c", "\u030c": "\u030c",
  "\u02d8": "\u0306", "\u0306": "\u0306",
  "\u00b4": "\u0301", "`": "\u0300", "\u02da": "\u030a",
};

function name(node: Node) {
  return (node.localName || node.nodeName || "").replace(/^.*:/, "").toLowerCase();
}

function elements(node: Node): Node[] {
  return Array.from(node.childNodes).filter((child) => child.nodeType === ELEMENT);
}

function attribute(node: Node, key: string) {
  return node.getAttribute?.(key) ?? null;
}

/** One piece of linear text and whether it reads as a single unit (needs no
 *  parentheses as a fraction's part or an exponent). */
type Piece = { text: string; atomic: boolean; space?: boolean; binomial?: boolean; operator?: boolean };

class Encoder {
  notes = new Set<string>();

  piece(node: Node): Piece {
    const tag = name(node);
    switch (tag) {
      case "mstyle":
        if (attribute(node, "mathcolor") || attribute(node, "mathbackground")) this.notes.add("Colours are dropped.");
        return this.row(elements(node));
      case "menclose":
        this.notes.add("A box, a strike or a circle around a part is dropped.");
        return this.row(elements(node));
      case "mphantom":
        this.notes.add("Invisible spacing (\\phantom) is dropped.");
        return { text: "", atomic: true };
      case "mglyph":
        this.notes.add("A custom symbol has no Unicode character.");
        return { text: "", atomic: true };
      case "math":
      case "mrow":
      case "mpadded":
      case "semantics":
      case "mtd":
        return this.row(tag === "semantics" ? elements(node).slice(0, 1) : elements(node));
      case "annotation":
      case "annotation-xml":
        return { text: "", atomic: true };
      case "mi":
      case "mn":
      case "mtext":
      case "ms":
        return this.token(node);
      case "mo":
        return { text: this.operator(node), atomic: true };
      case "mspace": {
        const text = this.space(node);
        return { text, atomic: true, space: Boolean(text) };
      }
      case "msup":
        return this.scripts(elements(node)[0], undefined, elements(node)[1]);
      case "msub":
        return this.scripts(elements(node)[0], elements(node)[1], undefined);
      case "msubsup":
        return this.scripts(elements(node)[0], elements(node)[1], elements(node)[2]);
      case "munder":
        return this.underOver(node, elements(node)[0], elements(node)[1], undefined);
      case "mover":
        return this.underOver(node, elements(node)[0], undefined, elements(node)[1]);
      case "munderover":
        return this.underOver(node, elements(node)[0], elements(node)[1], elements(node)[2]);
      case "mfrac":
        return this.fraction(node);
      case "msqrt":
        return this.radical(this.row(elements(node)), undefined);
      case "mroot":
        return this.radical(this.piece(elements(node)[0]), this.piece(elements(node)[1]));
      case "mfenced":
        return this.fenced(node);
      case "mtable":
        return this.table(node);
      case "mtr":
      case "mlabeledtr":
        return { text: elements(node).map((cell) => this.piece(cell).text.trim()).join(", "), atomic: false };
      case "merror":
        this.notes.add("A part of the formula could not be converted.");
        return { text: node.textContent ?? "", atomic: false };
      default:
        return this.row(elements(node));
    }
  }

  /** A row: its pieces joined with spaces around binary operators and
   *  relations, as text writes them. */
  row(children: Node[]): Piece {
    let binomial: Piece | null = null;
    const parts: { text: string; kind: "operand" | "binary" | "relation" | "open" | "close" | "comma" | "space" | "function" | "large"; atomic?: boolean }[] = [];
    for (const original of children) {
      // a fence or an operator alone in a group (\left(, \binom's parentheses)
      let child = original;
      while ((name(child) === "mrow" || name(child) === "mstyle") && elements(child).length === 1) child = elements(child)[0];
      const tag = name(child);
      if (tag === "mo") {
        const text = this.operator(child);
        if (!text) {
          // function application: "sin x"
          if ((child.textContent ?? "") === "\u2061") parts.push({ text: "", kind: "function" });
          continue;
        }
        if (text === "," || text === ";") parts.push({ text, kind: "comma" });
        else if (RELATIONS.has(text)) parts.push({ text, kind: "relation" });
        else if (BINARY.has(text)) parts.push({ text, kind: "binary" });
        else if (OPENING.has(text) && attribute(child, "form") !== "postfix") parts.push({ text, kind: "open" });
        else if (")]}⟩⌋⌉".includes(text)) parts.push({ text, kind: "close" });
        else parts.push({ text, kind: LARGE_OPERATORS.test(text) || isNamedOperator(child) ? "large" : "operand" });
        continue;
      }
      const piece = this.piece(child);
      if (piece.space) {
        parts.push({ text: " ", kind: "space" });
        continue;
      }
      if (piece.binomial) binomial = piece;
      if (piece.text) parts.push({ text: piece.text, kind: LARGE_OPERATORS.test(piece.text) || piece.operator ? "large" : "operand", atomic: piece.atomic });
    }
    const meaningful = parts.filter((part) => part.kind !== "space" && part.kind !== "function");
    if (!meaningful.length) return { text: "", atomic: true, space: parts.some((part) => part.kind === "space") };
    // \binom: its parentheses are part of C(n, k).
    if (binomial && meaningful.length === 3 && meaningful[0].kind === "open" && meaningful[2].kind === "close") {
      return { text: binomial.text, atomic: true };
    }
    let text = "";
    let previous: (typeof parts)[number]["kind"] | null = null;
    for (const part of parts) {
      if (part.kind === "function") {
        previous = "function";
        continue;
      }
      if (part.kind === "binary" || part.kind === "relation") {
        // A sign at the start, after an operator or an opening: unary, no spaces.
        const unary = previous === null || previous === "binary" || previous === "relation" || previous === "open" || previous === "comma";
        if (unary && part.kind === "binary") text += part.text;
        else text = text.replace(/ $/, "") + ` ${part.text} `;
      } else if (part.kind === "comma") {
        text = text.replace(/ $/, "") + `${part.text} `;
      } else if (part.kind === "space") {
        if (text && !text.endsWith(" ")) text += " ";
      } else if (part.kind === "close") {
        text = text.replace(/ $/, "") + part.text;
      } else {
        // "sin x", "lim_(x → 0) (sin x)/x", "∫ₐᵇ f(x)": a space after a
        // function name or a large operator.
        const opensParen = part.kind === "open" || (part.atomic === true && /^[(\[]/.test(part.text));
        const spaced =
          (previous === "function" && !opensParen) ||
          (previous === "large" && !(opensParen && /[A-Za-z]$/.test(text)));
        if (spaced && text && !text.endsWith(" ")) text += " ";
        text += part.text;
      }
      previous = part.kind;
    }
    text = text.trim();
    const atomic =
      meaningful.length <= 1 ||
      (meaningful[0]?.kind === "open" && meaningful.at(-1)?.kind === "close" && closesAtEnd(text)) ||
      // a differential: ∂f/∂x, dy/dx
      (meaningful.length === 2 && /^(∂|d)$/.test(meaningful[0].text) && Array.from(meaningful[1].text).length === 1);
    return { text, atomic };
  }

  token(node: Node): Piece {
    const raw = (node.textContent ?? "").replace(/\s+/g, " ");
    const variant = attribute(node, "mathvariant");
    const result = variant && ALPHABETS[variant] ? styled(raw, variant) : { text: raw, complete: true };
    const text = result.text;
    if (!result.complete) this.notes.add(`Some letters have no ${variant} Unicode form and stay plain.`);
    // A letter style Unicode has no alphabet for (sans-serif italic, ...).
    if (variant && !ALPHABETS[variant] && variant !== "normal" && variant !== "italic" && /[A-Za-z0-9]/.test(raw)) {
      this.notes.add(`The ${variant} letter style has no Unicode characters; plain letters are used.`);
    }
    const tag = name(node);
    if (tag === "mtext") return { text, atomic: !/\s/.test(text.trim()) };
    return { text, atomic: tag === "mi" || /^[\d.,]+$/.test(text) || Array.from(text).length <= 1 };
  }

  operator(node: Node): string {
    const text = (node.textContent ?? "").trim();
    if (INVISIBLE.has(text)) return "";
    return text === "-" ? "−" : text;
  }

  space(node: Node): string {
    const width = attribute(node, "width") ?? "";
    const value = parseFloat(width);
    if (/em$/.test(width) && value >= 0.15) return " ";
    if (attribute(node, "linebreak") === "newline") return " ";
    return "";
  }

  scripts(baseNode?: Node, subNode?: Node, supNode?: Node): Piece {
    const base = baseNode ? this.piece(baseNode) : { text: "", atomic: true };
    // F(x)|ₐᵇ: what ends in a closing delimiter carries its scripts as it is.
    const closed = /[)\]}|‖⟩]$/.test(base.text) && !/\s/.test(base.text);
    let text = base.atomic || closed || !base.text ? base.text : `(${base.text})`;
    if (subNode) text += this.script(this.piece(subNode), SUBSCRIPTS, "_");
    if (supNode) text += this.script(this.piece(supNode), SUPERSCRIPTS, "^");
    return { text, atomic: true, operator: Boolean(baseNode && name(baseNode) === "mo" && isNamedOperator(baseNode)) || LARGE_OPERATORS.test(text) };
  }

  /** A script as script characters when all of it has them and it is short
   *  and simple (x², x⁻¹, aᵢⱼ, xⁿ⁺¹), else ^x or ^(…) (x^(a + b)). */
  script(piece: Piece, table: Record<string, string>, mark: string) {
    const compact = piece.text.replace(/\s+/g, "");
    if (compact === "′" || compact === "″" || compact === "‴") return compact;
    // a styled letter maps through its plain letter (𝖳 -> ᵀ); a script inside
    // the script (x² in e^{-x²}) does not flatten
    const chars = Array.from(compact).map((char) =>
      char in table || !/[\u{1d400}-\u{1d7ff}]/u.test(char) ? char : char.normalize("NFKC"),
    );
    const operators = chars.filter((char) => "+-−=".includes(char)).length;
    const letters = chars.filter((char) => /\p{L}/u.test(char)).length;
    const signedNumber = /^[-−+]?[0-9]+$/.test(compact);
    const simple =
      signedNumber ||
      (operators === 0 && chars.length <= 3) ||
      (operators === 1 && letters <= 1 && chars.length <= 4 && !"+-−=".includes(chars.at(-1) ?? ""));
    if (chars.length && simple && chars.every((char) => char in table)) return chars.map((char) => table[char]).join("");
    this.notes.add("Scripts without Unicode script characters are written with ^ and _.");
    if (piece.atomic && chars.length === 1) return `${mark}${compact}`;
    return `${mark}(${piece.text})`;
  }

  underOver(node: Node, baseNode?: Node, underNode?: Node, overNode?: Node): Piece {
    const base = baseNode ? this.piece(baseNode) : { text: "", atomic: true };
    // An accent over (or under) the base: combining marks on one character.
    const accentNode = overNode ?? underNode;
    const marks = overNode ? ACCENTS : UNDER_ACCENTS;
    const isAccent = attribute(node, overNode ? "accent" : "accentunder") === "true" ||
      (accentNode && name(accentNode) === "mo" && (accentNode.textContent ?? "").trim() in marks && !underNode !== !overNode);
    if (isAccent && accentNode && !(underNode && overNode)) {
      const mark = marks[(accentNode.textContent ?? "").trim()];
      if (mark && Array.from(base.text).length === 1) return { text: base.text + mark, atomic: true };
      if (mark) {
        this.notes.add("An accent over several characters is written over each of them.");
        return { text: Array.from(base.text).map((char) => (char === " " ? char : char + mark)).join(""), atomic: true };
      }
    }
    const operator = Boolean(baseNode && name(baseNode) === "mo" && isNamedOperator(baseNode)) || LARGE_OPERATORS.test(base.text);
    if (operator) {
      // Limits under and over an operator: as scripts after it (lim, max...
      // with a limit under them are written so in text too).
      if (overNode || !/^(lim|max|min|sup|inf)/.test(base.text)) this.notes.add("Limits are written as scripts after their operator.");
    } else {
      this.notes.add("What stands over or under a symbol (\\overset, \\underbrace...) is written as a script after it.");
    }
    return this.scripts(baseNode, underNode, overNode);
  }

  fraction(node: Node): Piece {
    const [numeratorNode, denominatorNode] = elements(node);
    const numerator = numeratorNode ? this.piece(numeratorNode) : { text: "", atomic: true };
    const denominator = denominatorNode ? this.piece(denominatorNode) : { text: "", atomic: true };
    if (attribute(node, "linethickness") === "0" || attribute(node, "linethickness") === "0px") {
      this.notes.add("A binomial coefficient is written as C(n, k).");
      return { text: `C(${numerator.text}, ${denominator.text})`, atomic: true, binomial: true };
    }
    this.notes.add("Fractions are written with a slash.");
    const wrap = (piece: Piece) => (piece.atomic ? piece.text : `(${piece.text})`);
    return { text: `${wrap(numerator)}/${wrap(denominator)}`, atomic: false };
  }

  radical(body: Piece, index: Piece | undefined): Piece {
    const inner = body.atomic ? body.text : `(${body.text})`;
    if (!index) return { text: `√${inner}`, atomic: true };
    const degree = index.text.replace(/\s+/g, "");
    if (degree === "3") return { text: `∛${inner}`, atomic: true };
    if (degree === "4") return { text: `∜${inner}`, atomic: true };
    const chars = Array.from(degree);
    if (chars.every((char) => char in SUPERSCRIPTS)) {
      return { text: `${chars.map((char) => SUPERSCRIPTS[char]).join("")}√${inner}`, atomic: true };
    }
    return { text: `root(${index.text}, ${body.text})`, atomic: true };
  }

  fenced(node: Node): Piece {
    const open = attribute(node, "open") ?? "(";
    const close = attribute(node, "close") ?? ")";
    const separator = (attribute(node, "separators") ?? ",").trim().charAt(0) || ",";
    const inner = elements(node).map((child) => this.piece(child).text).join(`${separator} `);
    return { text: `${open}${inner}${close}`, atomic: true };
  }

  table(node: Node): Piece {
    this.notes.add("Matrices are written row by row, rows separated by semicolons.");
    const rows = elements(node).map((row) => this.piece(row).text);
    return { text: rows.join("; "), atomic: false };
  }
}

/** \lim, \max, \sum...: what follows it is its argument, after a space. */
function isNamedOperator(node: Node) {
  return attribute(node, "data-mjx-texclass") === "OP" || attribute(node, "movablelimits") === "true" ||
    /^[A-Za-z]{2,}$/.test((node.textContent ?? "").trim());
}

function closesAtEnd(text: string) {
  // "(a + b)" is one unit; "(a) + (b)" is not.
  let depth = 0;
  const chars = Array.from(text);
  for (let index = 0; index < chars.length; index += 1) {
    const char = chars[index];
    if ("([{⟨⌊⌈".includes(char)) depth += 1;
    else if (")]}⟩⌋⌉".includes(char)) {
      depth -= 1;
      if (depth === 0 && index < chars.length - 1) return false;
    }
  }
  return depth === 0;
}

/** Letters and digits in a mathematical alphabet; `complete` is false when
 *  one of them has no character there. */
function styled(text: string, variant: string) {
  const alphabet = ALPHABETS[variant];
  let complete = true;
  const result = Array.from(text)
    .map((char) => {
      if (alphabet.exceptions?.[char]) return alphabet.exceptions[char];
      const code = char.codePointAt(0)!;
      if (code >= 65 && code <= 90 && alphabet.upper) return String.fromCodePoint(alphabet.upper + code - 65);
      if (code >= 97 && code <= 122 && alphabet.lower) return String.fromCodePoint(alphabet.lower + code - 97);
      if (code >= 48 && code <= 57 && alphabet.digit) return String.fromCodePoint(alphabet.digit + code - 48);
      if (code >= 0x391 && code <= 0x3a9 && code !== 0x3a2 && alphabet.greekUpper) return String.fromCodePoint(alphabet.greekUpper + code - 0x391);
      if (code >= 0x3b1 && code <= 0x3c9 && alphabet.greekLower) return String.fromCodePoint(alphabet.greekLower + code - 0x3b1);
      if (/[\p{L}\p{N}]/u.test(char)) complete = false;
      return char;
    })
    .join("");
  return { text: result, complete };
}

/** Unicode text for a formula's MathML. `parse` is DOMParser's in the page
 *  (a test passes another XML parser). */
export function mathMlToUnicode(
  mathMl: string,
  parse: (xml: string) => { documentElement: Node | null } = (xml) =>
    new DOMParser().parseFromString(xml, "application/xml") as unknown as { documentElement: Node | null },
): UnicodeMath {
  const root = parse(mathMl).documentElement;
  if (!root || name(root) === "parsererror") return { text: "", notes: ["The formula could not be read."] };
  const encoder = new Encoder();
  const text = encoder
    .piece(root)
    .text.replace(/[ \u00a0]{2,}/g, " ")
    .trim();
  return { text, notes: [...encoder.notes] };
}
