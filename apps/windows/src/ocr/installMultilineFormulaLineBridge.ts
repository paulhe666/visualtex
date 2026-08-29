import { useEditorStore } from "../stores/editorStore.ts";
import {
  clearPendingMultilineOcrFormula,
  peekPendingMultilineOcrFormula,
} from "./multilineFormula.ts";
import {
  expandFormulaLinesForPendingOcr,
  type FormulaLineLike,
} from "./multilineFormulaLineExpansion.ts";

type EditorStateLike = {
  lines?: FormulaLineLike[];
};

type StoreLike = {
  subscribe: (listener: (state: EditorStateLike) => void) => () => void;
  setState: (partial: Partial<EditorStateLike>) => void;
};

const installationKey = "__visualtexMultilineOcrLineBridgeInstalled";
type BridgeGlobal = typeof globalThis & { [installationKey]?: boolean };

/**
 * OCR recognition still flows through the existing editor insertion APIs. This
 * bridge observes only the exact pending OCR payload and expands that one changed
 * FormulaLine into independently editable lines. It never rewrites an unrelated
 * hand-authored align/matrix expression or a document loaded from disk.
 */
export function installMultilineFormulaLineBridge(): void {
  const target = globalThis as BridgeGlobal;
  if (target[installationKey]) return;
  target[installationKey] = true;

  const store = useEditorStore as unknown as StoreLike;
  let applying = false;
  store.subscribe((state) => {
    if (applying || !Array.isArray(state.lines)) return;
    const pending = peekPendingMultilineOcrFormula();
    if (!pending) return;
    const expanded = expandFormulaLinesForPendingOcr(state.lines, pending);
    if (!expanded) return;
    applying = true;
    try {
      clearPendingMultilineOcrFormula();
      store.setState({ lines: expanded });
    } finally {
      applying = false;
    }
  });
}

installMultilineFormulaLineBridge();
