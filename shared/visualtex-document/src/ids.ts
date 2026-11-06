// Stable ids for blocks and formulas: given once, kept through edits and
// "turn into", renewed only for copies (a pasted or split node whose id is
// already taken).
import { Plugin, PluginKey, type Transaction } from "prosemirror-state";
import type { Node } from "prosemirror-model";

export const idsKey = new PluginKey("visualtex-document-ids");

let counter = 0;

/** A short id, unique within the session and practically across documents. */
export function newNodeId(): string {
  counter = (counter + 1) % 1296;
  const time = Date.now().toString(36).slice(-5);
  const random = Math.floor(Math.random() * 46656).toString(36).padStart(3, "0");
  return time + random + counter.toString(36).padStart(2, "0");
}

function hasIdAttr(node: Node) {
  return Object.prototype.hasOwnProperty.call(node.type.spec.attrs ?? {}, "id");
}

/** The transaction that gives every id-carrying node a unique id, or null when all have one. */
export function assignIds(doc: Node, tr: Transaction | null, makeTr: () => Transaction): Transaction | null {
  const seen = new Set<string>();
  let result = tr;
  doc.descendants((node, pos) => {
    if (!hasIdAttr(node)) return;
    const current = node.attrs.id as string | null;
    if (current && !seen.has(current)) {
      seen.add(current);
      return;
    }
    let fresh = newNodeId();
    while (seen.has(fresh)) fresh = newNodeId();
    seen.add(fresh);
    result ??= makeTr();
    result.setNodeAttribute(pos, "id", fresh);
  });
  return result;
}

export function idsPlugin(): Plugin {
  return new Plugin({
    key: idsKey,
    appendTransaction(transactions, _old, state) {
      if (!transactions.some((tr) => tr.docChanged)) return null;
      // prosemirror-history keeps an appended transaction in the event of the one it follows.
      return assignIds(state.doc, null, () => state.tr);
    },
  });
}
