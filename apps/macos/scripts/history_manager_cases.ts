import assert from "node:assert/strict";
import {
  EDIT_GROUP_TIMEOUT_MS,
  HistoryManager,
} from "../src/history/HistoryManager";
import type {
  DocumentSnapshot,
  HistoryEntry,
  MathSelectionSnapshot,
  ReplayDirection,
} from "../src/history/historyTypes";
import type { FormulaLine } from "../src/types/formula";

const selection = (position: number): MathSelectionSnapshot => ({
  ranges: [[position, position]],
  direction: "none",
});

interface MockDocument {
  title: string;
  lines: FormulaLine[];
  activeLineId: string | null;
  selectionByLineId: Record<string, MathSelectionSnapshot>;
}

function cloneDocument(document: MockDocument): DocumentSnapshot {
  return {
    title: document.title,
    lines: document.lines.map((line) => ({ ...line })),
    activeLineId: document.activeLineId,
    selectionByLineId: Object.fromEntries(
      Object.entries(document.selectionByLineId).map(([lineId, value]) => [
        lineId,
        {
          ranges: value.ranges.map(([start, end]) => [start, end]),
          direction: value.direction,
        },
      ]),
    ),
  };
}

function createHarness(initialLines: FormulaLine[] = [{ id: "line-1", latex: "" }]) {
  const document: MockDocument = {
    title: "Document",
    lines: initialLines.map((line) => ({ ...line })),
    activeLineId: initialLines[0]?.id ?? null,
    selectionByLineId: Object.fromEntries(
      initialLines.map((line) => [line.id, selection(line.latex.length)]),
    ),
  };
  const applied: Array<{ entry: HistoryEntry; direction: ReplayDirection }> = [];
  const manager = new HistoryManager();

  const apply = (entry: HistoryEntry, direction: ReplayDirection) => {
    applied.push({ entry, direction });
    const undoing = direction === "undo";
    switch (entry.type) {
      case "replace-formula": {
        const line = document.lines.find((item) => item.id === entry.lineId);
        assert.ok(line, `missing line ${entry.lineId}`);
        line.latex = undoing ? entry.beforeLatex : entry.afterLatex;
        document.activeLineId = undoing
          ? entry.beforeActiveLineId
          : entry.afterActiveLineId;
        document.selectionByLineId[entry.lineId] = undoing
          ? entry.beforeSelection
          : entry.afterSelection;
        break;
      }
      case "add-line":
        if (undoing) {
          document.lines = document.lines.filter((line) => line.id !== entry.line.id);
          document.activeLineId = entry.beforeActiveLineId;
        } else {
          document.lines.splice(entry.index, 0, { ...entry.line });
          document.activeLineId = entry.afterActiveLineId;
        }
        break;
      case "remove-line":
        if (undoing) {
          document.lines.splice(entry.index, 0, { ...entry.line });
          document.activeLineId = entry.beforeActiveLineId;
        } else {
          document.lines = document.lines.filter((line) => line.id !== entry.line.id);
          document.activeLineId = entry.afterActiveLineId;
        }
        break;
      case "replace-document": {
        const snapshot = undoing ? entry.before : entry.after;
        document.title = snapshot.title;
        document.lines = snapshot.lines.map((line) => ({ ...line }));
        document.activeLineId = snapshot.activeLineId;
        document.selectionByLineId = { ...snapshot.selectionByLineId };
        break;
      }
      case "change-title":
        document.title = undoing ? entry.beforeTitle : entry.afterTitle;
        break;
    }
  };

  manager.configure({
    applyEntry: apply,
    getDocumentSnapshot: () => cloneDocument(document),
  });
  return { manager, document, applied };
}

function recordFormula(
  manager: HistoryManager,
  input: {
    lineId?: string;
    before: string;
    after: string;
    beforePosition: number;
    afterPosition: number;
    timestamp: number;
    editKind?: "insert" | "delete-backward" | "delete-forward" | "composition" | "replace";
    source?: "keyboard" | "toolbar" | "candidate" | "ocr" | "paste";
  },
) {
  manager.recordFormulaEdit({
    lineId: input.lineId ?? "line-1",
    beforeLatex: input.before,
    afterLatex: input.after,
    beforeSelection: selection(input.beforePosition),
    afterSelection: selection(input.afterPosition),
    beforeActiveLineId: input.lineId ?? "line-1",
    afterActiveLineId: input.lineId ?? "line-1",
    editKind: input.editKind ?? "insert",
    source: input.source ?? "keyboard",
    timestamp: input.timestamp,
  });
}

async function run() {
  {
    const { manager, document } = createHarness([{ id: "line-1", latex: "" }]);
    recordFormula(manager, {
      before: "",
      after: "a",
      beforePosition: 0,
      afterPosition: 1,
      timestamp: 0,
    });
    recordFormula(manager, {
      before: "a",
      after: "ab",
      beforePosition: 1,
      afterPosition: 2,
      timestamp: 100,
    });
    recordFormula(manager, {
      before: "ab",
      after: "abc",
      beforePosition: 2,
      afterPosition: 3,
      timestamp: 200,
    });
    manager.commitPendingTransaction();
    assert.equal(manager.getState().undoStack.length, 1, "continuous insert should merge");
    document.lines[0].latex = "abc";
    await manager.undo();
    assert.equal(document.lines[0].latex, "");
    await manager.redo();
    assert.equal(document.lines[0].latex, "abc");
  }

  {
    const { manager } = createHarness();
    recordFormula(manager, {
      before: "",
      after: "a",
      beforePosition: 0,
      afterPosition: 1,
      timestamp: 0,
    });
    recordFormula(manager, {
      before: "a",
      after: "ab",
      beforePosition: 1,
      afterPosition: 2,
      timestamp: EDIT_GROUP_TIMEOUT_MS + 1,
    });
    manager.commitPendingTransaction();
    assert.equal(manager.getState().undoStack.length, 2, "timeout should split edits");
  }

  {
    const { manager } = createHarness();
    recordFormula(manager, {
      before: "",
      after: "a",
      beforePosition: 0,
      afterPosition: 1,
      timestamp: 0,
    });
    recordFormula(manager, {
      before: "a",
      after: "ba",
      beforePosition: 0,
      afterPosition: 1,
      timestamp: 100,
    });
    manager.commitPendingTransaction();
    assert.equal(manager.getState().undoStack.length, 2, "selection jump must split edits");
  }

  {
    const { manager } = createHarness([{ id: "line-1", latex: "abc" }]);
    recordFormula(manager, {
      before: "abc",
      after: "ab",
      beforePosition: 3,
      afterPosition: 2,
      timestamp: 0,
      editKind: "delete-backward",
    });
    recordFormula(manager, {
      before: "ab",
      after: "a",
      beforePosition: 2,
      afterPosition: 1,
      timestamp: 100,
      editKind: "delete-backward",
    });
    manager.commitPendingTransaction();
    assert.equal(manager.getState().undoStack.length, 1, "continuous backspace should merge");
  }

  {
    const { manager } = createHarness([{ id: "line-1", latex: "abc" }]);
    recordFormula(manager, {
      before: "abc",
      after: "bc",
      beforePosition: 0,
      afterPosition: 0,
      timestamp: 0,
      editKind: "delete-forward",
    });
    recordFormula(manager, {
      before: "bc",
      after: "c",
      beforePosition: 0,
      afterPosition: 0,
      timestamp: 100,
      editKind: "delete-forward",
    });
    manager.commitPendingTransaction();
    assert.equal(manager.getState().undoStack.length, 1, "continuous Delete should merge");
  }

  {
    const { manager } = createHarness();
    manager.recordTitleEdit({
      beforeTitle: "D",
      afterTitle: "Do",
      timestamp: 0,
    });
    manager.recordTitleEdit({
      beforeTitle: "Do",
      afterTitle: "Doc",
      timestamp: 100,
    });
    manager.commitPendingTransaction();
    assert.equal(manager.getState().undoStack.length, 1, "title typing should merge");
    assert.equal(manager.getState().undoStack[0]?.type, "change-title");
  }

  {
    const { manager } = createHarness();
    recordFormula(manager, {
      before: "",
      after: "中文",
      beforePosition: 0,
      afterPosition: 2,
      timestamp: 0,
      editKind: "composition",
    });
    assert.equal(manager.getState().pendingTransaction, null);
    assert.equal(manager.getState().undoStack.length, 1, "composition should commit once");
  }

  {
    const { manager, document, applied } = createHarness([
      { id: "line-1", latex: "a" },
      { id: "line-2", latex: "b" },
      { id: "line-3", latex: "c" },
    ]);
    recordFormula(manager, {
      lineId: "line-1",
      before: "a",
      after: "a_1",
      beforePosition: 1,
      afterPosition: 3,
      timestamp: 0,
    });
    manager.commitPendingTransaction();
    recordFormula(manager, {
      lineId: "line-3",
      before: "c",
      after: "c_3",
      beforePosition: 1,
      afterPosition: 3,
      timestamp: 100,
    });
    manager.commitPendingTransaction();
    document.lines[0].latex = "a_1";
    document.lines[2].latex = "c_3";
    document.activeLineId = "line-2";
    await manager.undo();
    assert.equal(applied.at(-1)?.entry.type, "replace-formula");
    assert.equal((applied.at(-1)?.entry as { lineId: string }).lineId, "line-3");
    assert.equal(document.lines[2].latex, "c");
    await manager.undo();
    assert.equal((applied.at(-1)?.entry as { lineId: string }).lineId, "line-1");
  }

  {
    const { manager, document } = createHarness([{ id: "line-1", latex: "a" }]);
    const added = { id: "stable-added", latex: "" };
    manager.push({
      type: "add-line",
      line: added,
      index: 1,
      beforeActiveLineId: "line-1",
      afterActiveLineId: added.id,
      beforeSelection: selection(1),
      afterSelection: selection(0),
      timestamp: 1,
    });
    document.lines.push({ ...added });
    document.activeLineId = added.id;
    await manager.undo();
    assert.deepEqual(document.lines.map((line) => line.id), ["line-1"]);
    await manager.redo();
    assert.deepEqual(document.lines.map((line) => line.id), ["line-1", "stable-added"]);
  }

  {
    const { manager, document } = createHarness([
      { id: "line-1", latex: "a" },
      { id: "stable-removed", latex: "b" },
      { id: "line-3", latex: "c" },
    ]);
    const removed = { ...document.lines[1] };
    manager.push({
      type: "remove-line",
      line: removed,
      index: 1,
      beforeActiveLineId: removed.id,
      afterActiveLineId: "line-1",
      beforeSelection: selection(1),
      afterSelection: selection(1),
      timestamp: 1,
    });
    document.lines.splice(1, 1);
    await manager.undo();
    assert.equal(document.lines[1].id, "stable-removed");
    await manager.redo();
    assert.ok(!document.lines.some((line) => line.id === "stable-removed"));
  }

  {
    const { manager, document } = createHarness([{ id: "line-1", latex: "a" }]);
    const before = cloneDocument(document);
    const after: DocumentSnapshot = {
      title: "Source",
      lines: [
        { id: "line-1", latex: "x" },
        { id: "line-2", latex: "y" },
      ],
      activeLineId: "line-2",
      selectionByLineId: { "line-2": selection(1) },
    };
    manager.push({
      type: "replace-document",
      before,
      after,
      source: "source-apply",
      timestamp: 1,
    });
    Object.assign(document, {
      title: after.title,
      lines: after.lines.map((line) => ({ ...line })),
      activeLineId: after.activeLineId,
      selectionByLineId: after.selectionByLineId,
    });
    await manager.undo();
    assert.equal(document.lines[0].latex, "a");
    await manager.redo();
    assert.deepEqual(document.lines.map((line) => line.latex), ["x", "y"]);
  }

  {
    const { manager, document } = createHarness([{ id: "line-1", latex: "a" }]);
    manager.push({
      type: "replace-formula",
      lineId: "line-1",
      beforeLatex: "a",
      afterLatex: "a\\theta",
      beforeSelection: selection(1),
      afterSelection: selection(7),
      beforeActiveLineId: "line-1",
      afterActiveLineId: "line-1",
      timestamp: 1,
      source: "ocr",
    });
    document.lines[0].latex = "a\\theta";
    document.selectionByLineId["line-1"] = selection(7);
    await manager.undo();
    assert.equal(document.lines[0].latex, "a", "OCR undo should restore full formula");
    assert.deepEqual(document.selectionByLineId["line-1"], selection(1));
    await manager.redo();
    assert.equal(document.lines[0].latex, "a\\theta", "OCR redo should restore result");
    assert.deepEqual(document.selectionByLineId["line-1"], selection(7));
  }

  {
    const { manager, document } = createHarness([{ id: "line-1", latex: "a" }]);
    const before = cloneDocument(document);
    const after: DocumentSnapshot = {
      title: "Document",
      lines: [
        { id: "line-1", latex: "a" },
        { id: "ocr-line-2", latex: "b" },
      ],
      activeLineId: "ocr-line-2",
      selectionByLineId: { "ocr-line-2": selection(1) },
    };
    manager.push({
      type: "replace-document",
      before,
      after,
      source: "ocr",
      timestamp: 1,
    });
    Object.assign(document, {
      lines: after.lines.map((line) => ({ ...line })),
      activeLineId: after.activeLineId,
      selectionByLineId: after.selectionByLineId,
    });
    assert.ok(
      manager.getState().checkpoints.length >= 1,
      "large OCR edit should create L3 checkpoint",
    );
    await manager.undo();
    assert.deepEqual(document.lines.map((line) => line.id), ["line-1"]);
    await manager.redo();
    assert.deepEqual(document.lines.map((line) => line.id), ["line-1", "ocr-line-2"]);
  }

  {
    const { manager } = createHarness();
    for (let index = 0; index < 325; index += 1) {
      manager.push({
        type: "change-title",
        beforeTitle: `Title ${index}`,
        afterTitle: `Title ${index + 1}`,
        timestamp: index,
      });
    }
    assert.ok(
      manager.getState().undoStack.length <= 300,
      "L2 history must respect the capacity limit",
    );
    assert.ok(
      manager.getState().checkpoints.length > 0,
      "capacity trimming should create an L3 checkpoint",
    );
  }

  {
    const { manager, document } = createHarness([{ id: "line-1", latex: "" }]);
    recordFormula(manager, {
      before: "",
      after: "a",
      beforePosition: 0,
      afterPosition: 1,
      timestamp: 0,
    });
    manager.commitPendingTransaction();
    document.lines[0].latex = "a";
    await manager.undo();
    assert.equal(manager.getState().canRedo, true);
    recordFormula(manager, {
      before: "",
      after: "b",
      beforePosition: 0,
      afterPosition: 1,
      timestamp: 100,
    });
    assert.equal(
      manager.getState().redoStack.length,
      0,
      "new edit must clear redo immediately",
    );
    assert.equal(manager.getState().canRedo, false);
    manager.commitPendingTransaction();
  }

  {
    const { manager, document } = createHarness([{ id: "line-1", latex: "" }]);
    recordFormula(manager, {
      before: "",
      after: "a",
      beforePosition: 0,
      afterPosition: 1,
      timestamp: 0,
    });
    manager.commitPendingTransaction();
    document.lines[0].latex = "a";
    manager.configure({
      getDocumentSnapshot: () => cloneDocument(document),
      applyEntry: async (entry, direction) => {
        manager.push({
          type: "change-title",
          beforeTitle: "x",
          afterTitle: "y",
          timestamp: 99,
        });
        if (entry.type === "replace-formula") {
          document.lines[0].latex =
            direction === "undo" ? entry.beforeLatex : entry.afterLatex;
        }
      },
    });
    await manager.undo();
    assert.equal(manager.getState().undoStack.length, 0, "replay must not record recursively");
    assert.equal(manager.getState().redoStack.length, 1);
  }

  console.log("HistoryManager smoke test passed");
}

await run();
