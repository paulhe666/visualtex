import { useCallback, useRef, useState } from "react";
import type { CustomSymbolDesignerDocument } from "../designerTypes.ts";

const LIMIT = 200;

interface HistoryState {
  past: CustomSymbolDesignerDocument[];
  present: CustomSymbolDesignerDocument;
  future: CustomSymbolDesignerDocument[];
}

export type DocumentUpdate =
  | CustomSymbolDesignerDocument
  | ((current: CustomSymbolDesignerDocument) => CustomSymbolDesignerDocument);

/**
 * The document with undo and redo. Changes that carry the same `gesture` key
 * one after another (a drag, typing in one field) are one step.
 */
export function useDesignerHistory(initial: () => CustomSymbolDesignerDocument) {
  const [state, setState] = useState<HistoryState>(() => ({ past: [], present: initial(), future: [] }));
  const lastGesture = useRef<string | null>(null);

  const change = useCallback((update: DocumentUpdate, gesture?: string) => {
    setState((current) => {
      const next = typeof update === "function" ? update(current.present) : update;
      if (next === current.present) return current;
      const merge = gesture !== undefined && gesture === lastGesture.current;
      lastGesture.current = gesture ?? null;
      return {
        past: merge ? current.past : [...current.past, current.present].slice(-LIMIT),
        present: next,
        future: [],
      };
    });
  }, []);

  const undo = useCallback(() => {
    lastGesture.current = null;
    setState((current) => {
      const previous = current.past[current.past.length - 1];
      if (!previous) return current;
      return { past: current.past.slice(0, -1), present: previous, future: [current.present, ...current.future] };
    });
  }, []);

  const redo = useCallback(() => {
    lastGesture.current = null;
    setState((current) => {
      const next = current.future[0];
      if (!next) return current;
      return { past: [...current.past, current.present], present: next, future: current.future.slice(1) };
    });
  }, []);

  /** A different document (opened, new): its own history. */
  const replace = useCallback((document: CustomSymbolDesignerDocument) => {
    lastGesture.current = null;
    setState({ past: [], present: document, future: [] });
  }, []);

  /** Ends the current gesture: the next change is a new step. */
  const endGesture = useCallback(() => {
    lastGesture.current = null;
  }, []);

  return {
    document: state.present,
    change,
    undo,
    redo,
    replace,
    endGesture,
    canUndo: state.past.length > 0,
    canRedo: state.future.length > 0,
  };
}
