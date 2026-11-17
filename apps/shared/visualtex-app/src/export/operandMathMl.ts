import type { MmlNode } from "mathjax-full/js/core/MmlTree/MmlNode.js";
import type { MmlFactory } from "mathjax-full/js/core/MmlTree/MmlFactory.js";
import {
  isDifferentialSymbol,
  isIntegralSymbol,
  operandScope,
  operandTokenOfOperator,
  type OperandToken,
} from "../../../tex-layout/src/operand-scope.ts";

// The Word writers (WordOmmlFromMathMl.NormalizeNaryArguments, C#, and its Rust port on macOS) make a native n-ary object of a
// big operator, and the operand it holds as its body is the row that follows the operator in the MathML; MathJax writes the
// items loose, and the writers then take only the single element after the operator (the f of `∫ f(x) dx`, the rest of the
// integrand outside the integral). The operand is found here, by the one rule of the editor's operand region
// (tex-layout/src/operand-scope.ts: an integral up to a sign after its differential, a sum up to a top-level sign, both up to a
// relation or a separator or the end of the row), and wrapped in an mrow, which the writers take whole.

/** The n-ary operators of Word's equations (the writers' NaryCharacters). */
const NARY_CHARACTERS = new Set(Array.from("∑∏∐∫∬∭∮∯∰∱∲∳⨋⨌⨍⨎⨏⨐⨑⨒⨓⨔⨕⨖⨗⨘⨙⨚⨛⨜⋀⋁⋂⋃"));
const LIMIT_KINDS = new Set(["munder", "mover", "munderover", "msub", "msup", "msubsup"]);
/** The elements whose children are one row of items (the writers' NaryRowParents): not msubsup and the like, whose children are a base and its scripts. */
const ROW_KINDS = new Set(["math", "mrow", "inferredMrow", "mstyle", "msqrt", "mtd", "mpadded", "menclose", "mphantom", "merror"]);

const textOf = (node: MmlNode) => ((node as unknown as { getText?: () => string }).getText?.() ?? "").trim();
const children = (node: MmlNode) => (node.childNodes ?? []) as MmlNode[];

/** The operator character of a row item that is a big operator (a bare mo, or one carried by a limit structure), else "". */
function naryOf(node: MmlNode): string {
  const carrier = node.kind === "mo" ? node : LIMIT_KINDS.has(node.kind) ? children(node)[0] : undefined;
  if (!carrier || carrier.kind !== "mo") return "";
  const text = textOf(carrier);
  return text.length > 0 && Array.from(text).every((ch) => NARY_CHARACTERS.has(ch)) ? text : "";
}

function tokenOf(node: MmlNode, next: MmlNode | undefined): OperandToken {
  if (node.kind === "mspace") return "space";
  // MathJax writes \, as an mstyle holding the mspace
  if (node.kind === "mstyle" && children(node).length === 1 && children(node)[0].kind === "mspace") return "space";
  const text = textOf(node);
  if ((node.kind === "mi" || node.kind === "mo") && isDifferentialSymbol(text)) return next?.kind === "mi" ? "differential" : "other";
  return node.kind === "mo" ? operandTokenOfOperator(text) : "other";
}

/**
 * Wraps the operand of every n-ary operator in the tree in an mrow (the items after the operator that it applies to). An operator
 * whose operand is empty, or already one mrow, is left alone.
 */
export function groupBigOperatorOperands(root: MmlNode, factory: MmlFactory): void {
  const rows: MmlNode[] = [];
  root.walkTree((node) => {
    if (ROW_KINDS.has((node as MmlNode).kind) && children(node as MmlNode).length > 1) rows.push(node as MmlNode);
  });
  for (const row of rows) {
    const items = children(row);
    const tokens = items.map((item, index) => tokenOf(item, items[index + 1]));
    const found: { op: MmlNode; from: number; to: number }[] = [];
    items.forEach((item, index) => {
      const symbol = naryOf(item);
      if (!symbol) return;
      const { end } = operandScope(tokens, index + 1, isIntegralSymbol(symbol));
      if (end > index + 1) found.push({ op: item, from: index + 1, to: end });
    });
    // an operator inside another's operand cannot reach past its end
    for (const outer of found) {
      for (const inner of found) {
        if (inner !== outer && inner.from > outer.from && inner.from <= outer.to && inner.to > outer.to) inner.to = Math.max(inner.from, outer.to);
      }
    }
    const spans = found
      .filter((g) => g.to > g.from)
      .map((g) => ({ first: items[g.from], last: items[g.to - 1] }));
    // the inner operands first (their operators stand later in the row), so an outer operand holds the inner one whole
    for (const span of spans.reverse()) {
      const siblings = children(row);
      const from = siblings.indexOf(span.first);
      let last: MmlNode | null = span.last;
      while (last && last.parent !== row) last = last.parent as MmlNode | null;
      const to = last ? siblings.indexOf(last) : -1;
      if (from < 0 || to < from) continue;
      // one item that already is a row stays as it is: the writers take it whole
      if (from === to && siblings[from].kind === "mrow" && !siblings[from].isInferred) continue;
      const group = factory.create("mrow") as MmlNode;
      const taken = siblings.splice(from, to - from + 1, group);
      group.parent = row;
      for (const node of taken) {
        node.parent = group;
        (group.childNodes as MmlNode[]).push(node);
      }
    }
  }
}
