// Where a big operator's operand ends (docs/editing-core/DESIGN.md §4, "operand"): ∫ ∑ ∏ and the rest of the n-ary
// operators apply to what follows them, and the formula source does not say how far. One rule, for everything that has to
// know: the editor's operand region (the L-shaped caret's underline under only the operand while the caret is in it, as
// AxMath draws it), the MathML the Word writers get (the operand a native Word n-ary object holds as its body), PowerPoint's
// OMML. It is the rule of the Word writer's loose operand (WordOmmlFromMathMl.NaryLooseOperand, ported in omml_write::prep):
// the operand runs up to a relation or a separator, to a fence that closes something opened before the operator, and to
// the end of the row; a sum also ends at a top-level sign (+ - ± ∓: `∑ aᵢ + b` is (∑ aᵢ) + b); an integral only at one that
// comes after its differential (`∫ x + 1 dx` holds the plus, `∫ f dx + c` does not).

/**
 * What an item of a row is to the rule: an opening or closing fence, a relation or separator (`stop`), a sign, the `d` of
 * a differential (followed by its variable), a row's last full stop, spacing (never the operand's last item), or any other.
 */
export type OperandToken = "open" | "close" | "stop" | "sign" | "differential" | "period" | "space" | "other";

const OPENERS = "([{⟨⌈⌊";
const CLOSERS = ")]}⟩⌉⌋";
const STOPS = "=<>≤≥≠≈≡∼≃≅∝→←↔⇒⇐⇔↦∈∉⊂⊆⊃⊇≪≫≺≻⩽⩾,;";
const SIGNS = "+-−±∓";

/** ∫ ∬ ∭ ∮ ∯ ∰ ∱ ∲ ∳ and the extended integrals of Unicode's Supplemental Mathematical Operators: an integral's operand ends after its differential. */
export function isIntegralSymbol(symbol: string): boolean {
  if (!symbol) return false;
  for (const ch of symbol) {
    const c = ch.codePointAt(0)!;
    if (!((c >= 0x222b && c <= 0x2233) || (c >= 0x2a0b && c <= 0x2a1c))) return false;
  }
  return true;
}

/** An operator token's class from its text (one character: a fence, a relation, a separator, a sign, a full stop). */
export function operandTokenOfOperator(text: string): OperandToken {
  if (text.length !== 1) return "other";
  const c = text[0];
  if (OPENERS.includes(c)) return "open";
  if (CLOSERS.includes(c)) return "close";
  if (STOPS.includes(c)) return "stop";
  if (SIGNS.includes(c)) return "sign";
  if (c === ".") return "period";
  return "other";
}

/** Whether a symbol is the differential d: `d` (math italic or upright) or ⅆ; the caller checks that a variable follows. */
export function isDifferentialSymbol(text: string): boolean {
  return text === "d" || text === "ⅆ";
}

export interface OperandScope {
  /** the operand is the items [from, end) of the row; end === from: it has no items */
  end: number;
  /** the index of the differential's `d` the operand holds, or -1 */
  differential: number;
}

/**
 * The operand of the operator before `from` (the row's items `tokens`, the first of the operand at `from`): how far it
 * runs. `integral`: the operator is an integral. The operand never ends in spacing.
 */
export function operandScope(tokens: readonly OperandToken[], from: number, integral: boolean): OperandScope {
  let depth = 0;
  let differential = -1;
  let i = from;
  for (; i < tokens.length; i++) {
    const token = tokens[i];
    if (token === "open") depth++;
    else if (token === "close") {
      if (depth === 0) break;
      depth--;
    } else if (depth === 0) {
      if (token === "stop") break;
      if (token === "period" && i === tokens.length - 1) break;
      // (an operand has at least one item before a sign ends it: the sign is then the operand's own unary sign)
      if (token === "sign" && i > from && (!integral || differential >= 0)) break;
      if (token === "differential" && differential < 0) differential = i;
    }
  }
  let end = i;
  while (end > from && tokens[end - 1] === "space") end--;
  return { end, differential: differential >= end ? -1 : differential };
}
