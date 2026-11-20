import {
  createContext, useContext, useEffect, useMemo, useState, useSyncExternalStore,
  type ReactNode, type RefObject,
} from "react";
import { flushSync } from "react-dom";
import { HistoryManager } from "./HistoryManager";
import { documentSnapshotsEquivalent } from "./documentSnapshot";
import { applyHistoryEntryToEditor, getEditorDocumentSnapshot } from "./documentHistory";
import type { FormulaDocument } from "../types/formula";
import { useEditorStore } from "../stores/editorStore";
import type { DocumentSnapshot, MathSelectionSnapshot, ReplaceDocumentEntry } from "./historyTypes";

const HistoryContext = createContext<HistoryManager | null>(null);

export function EditorSessionProvider({ children }: { children: ReactNode }) {
  const [history] = useState(() => new HistoryManager());
  return <HistoryContext.Provider value={history}>{children}</HistoryContext.Provider>;
}

export function useHistoryManager() {
  const history = useContext(HistoryContext);
  if (!history) throw new Error("An editor must belong to an EditorSessionProvider.");
  return history;
}

export function useHistorySnapshot() {
  const history = useHistoryManager();
  return useSyncExternalStore(history.subscribe, history.getSnapshot, history.getSnapshot);
}

interface SelectionController {
  captureDocumentSnapshot(): DocumentSnapshot;
  restoreDocumentSelection(snapshot: DocumentSnapshot): boolean;
  restoreSelection(
    lineId: string,
    latex: string,
    selection: MathSelectionSnapshot | null,
  ): Promise<boolean>;
}

// Desktop and Office share document transactions and history replay.
// The view restores selection after React has applied the document change.
export function useDocumentSession(editor: RefObject<SelectionController | null>) {
  const history = useHistoryManager();
  const session = useMemo(() => {
    const captureSnapshot = () =>
      editor.current?.captureDocumentSnapshot() ?? getEditorDocumentSnapshot();
    const restoreSnapshotFocus = (snapshot: DocumentSnapshot) => {
      editor.current?.restoreDocumentSelection(snapshot);
    };
    const replaceDocumentWithHistory = (
      after: DocumentSnapshot,
      source: ReplaceDocumentEntry["source"],
    ) => {
      if (source !== "source-apply") history.commitPendingTransaction();
      const before = captureSnapshot();
      if (documentSnapshotsEquivalent(before, after)) return false;
      const apply = () => useEditorStore.getState().replaceDocumentState(after);
      if (source === "source-apply") apply();
      else flushSync(apply);
      const entry: ReplaceDocumentEntry = {
        type: "replace-document", before, after, source, timestamp: Date.now(),
      };
      if (source === "source-apply") history.recordSourceDocumentEdit(entry);
      else {
        history.push(entry);
        restoreSnapshotFocus(after);
      }
      return true;
    };
    const openDocumentWithHistory = (document: FormulaDocument) => {
      history.commitPendingTransaction();
      const before = captureSnapshot();
      flushSync(() => useEditorStore.getState().loadDocument(document));
      const after = getEditorDocumentSnapshot();
      if (!documentSnapshotsEquivalent(before, after)) {
        history.push({
          type: "replace-document", before, after,
          source: "open-document", timestamp: Date.now(),
        });
        restoreSnapshotFocus(after);
      }
    };
    return { captureSnapshot, replaceDocumentWithHistory, openDocumentWithHistory };
  }, [editor, history]);

  useEffect(() => {
    history.configure({
      getDocumentSnapshot: session.captureSnapshot,
      applyEntry: async (entry, direction) => {
        const restored = flushSync(() => applyHistoryEntryToEditor(entry, direction));
        if (restored?.document) {
          editor.current?.restoreDocumentSelection(restored.document);
        } else if (restored) {
          await editor.current?.restoreSelection(
            restored.lineId, restored.latex, restored.selection,
          );
        }
      },
    });
    return () => {
      history.commitPendingTransaction();
      history.configure(null);
    };
  }, [editor, history, session]);
  return session;
}
