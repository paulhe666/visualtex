// The source tab in document mode: the document as Markdown, following the document as it is edited; what it
// shows is what Copy and Export write. It is also an
// editor: what is typed here is read into the document a moment after typing stops and when the editor is left
// (sourceEdit.ts: only the blocks that differ are replaced, one undo step in the document per change). While it is
// being edited the document does not write over it; once it is left it shows the document's source again.
// Unfinished code (an open formula or fence) is not read: the document keeps its last good state and the line is
// marked. What reading drops is listed in the header and its lines marked, before it goes when the editor is left.
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Annotation, Compartment, EditorState, StateEffect, StateField, Transaction, type Extension } from "@codemirror/state";
import { Decoration, EditorView, highlightActiveLineGutter, keymap, lineNumbers, type DecorationSet } from "@codemirror/view";
import { HighlightStyle, syntaxHighlighting } from "@codemirror/language";
import { defaultKeymap, history, historyKeymap, insertNewlineKeepIndent } from "@codemirror/commands";
import { tags } from "@lezer/highlight";
import type { EditorState as DocumentState } from "prosemirror-state";
import {
  documentToMarkdown,
  markdownLosses,
  markdownSourceOf,
  sourcePlaceOf,
  sourceProblem,
  sourceTransaction,
  type SourceLoss,
  type SourceProblem,
} from "../../shared/visualtex-document/src/index.ts";
import { useDocumentEditorStore } from "../stores/documentEditorStore";
import { useDocumentFormatStore } from "../stores/documentFormatStore";
import { useEditorStore } from "../stores/editorStore";
import type { Theme } from "../types/formula";
import { portableFormulaLatex } from "./DocumentEditor";

export const highlightStyle = HighlightStyle.define([
  { tag: [tags.keyword, tags.definitionKeyword, tags.macroName, tags.labelName, tags.heading], color: "var(--syntax-command)" },
  { tag: [tags.className, tags.typeName, tags.namespace], color: "var(--syntax-function)", fontWeight: "650" },
  { tag: [tags.strong, tags.emphasis, tags.monospace], color: "var(--syntax-command)", fontWeight: "600" },
  { tag: [tags.operator, tags.processingInstruction], color: "var(--syntax-operator)" },
  { tag: tags.number, color: "var(--syntax-number)" },
  { tag: tags.bracket, color: "var(--syntax-bracket)" },
  { tag: [tags.string, tags.quote, tags.meta], color: "var(--syntax-string)" },
  { tag: tags.comment, color: "var(--syntax-comment)", fontStyle: "italic" },
  { tag: [tags.variableName, tags.content], color: "var(--syntax-variable)" },
]);

export const theme: Extension = EditorView.theme({
  "&": { backgroundColor: "transparent", color: "var(--text)", height: "100%" },
  ".cm-scroller": { overflow: "auto" },
  ".cm-content": {
    fontFamily: "'SFMono-Regular', Menlo, Consolas, monospace",
    fontSize: "var(--visualtex-latex-source-font-size, 12px)",
    lineHeight: "1.62",
    padding: "10px 0 18px",
    caretColor: "var(--text)",
  },
  ".cm-line": { paddingInline: "4px 10px" },
  ".cm-gutters": { backgroundColor: "transparent", color: "var(--text-faint)", border: "none" },
  ".cm-lineNumbers .cm-gutterElement": { minWidth: "32px", paddingInline: "6px 7px" },
  ".cm-activeLineGutter": { backgroundColor: "transparent" },
  "&.cm-focused": { outline: "none" },
  ".cm-selectionBackground, ::selection": { backgroundColor: "color-mix(in srgb, var(--accent) 22%, transparent) !important" },
  // unfinished code (not read) and code reading drops
  ".cm-vtx-problem-line": { backgroundColor: "color-mix(in srgb, var(--danger) 10%, transparent)" },
  ".cm-vtx-problem": { textDecoration: "underline wavy var(--danger)", textUnderlineOffset: "3px" },
  ".cm-vtx-lost-line": { backgroundColor: "color-mix(in srgb, var(--warning) 9%, transparent)" },
});

/** Text put in from the document (not typed): not a change to read back. */
const fromDocument = Annotation.define<boolean>();

/** The lines marked: the unfinished piece, the lines with code that reading drops. */
const setMarks = StateEffect.define<{ problem: SourceProblem | null; lost: number[] }>();

const marksField = StateField.define<DecorationSet>({
  create: () => Decoration.none,
  update(marks, tr) {
    marks = marks.map(tr.changes);
    for (const effect of tr.effects) {
      if (!effect.is(setMarks)) continue;
      const { doc } = tr.state;
      const ranges = [];
      const lines = new Set(effect.value.lost.filter((n) => n >= 1 && n <= doc.lines));
      const problem = effect.value.problem;
      if (problem && problem.line >= 1 && problem.line <= doc.lines) {
        lines.delete(problem.line);
        ranges.push(Decoration.line({ class: "cm-vtx-problem-line" }).range(doc.line(problem.line).from));
        const from = Math.min(problem.from, doc.length);
        const to = Math.min(Math.max(problem.to, from + 1), doc.length);
        if (to > from) ranges.push(Decoration.mark({ class: "cm-vtx-problem" }).range(from, to));
      }
      for (const n of lines) ranges.push(Decoration.line({ class: "cm-vtx-lost-line" }).range(doc.line(n).from));
      marks = Decoration.set(ranges, true);
    }
    return marks;
  },
  provide: (field) => EditorView.decorations.from(field),
});

/** How long after typing stops what was typed is read into the document. */
const APPLY_DELAY_MS = 400;

/** The Markdown the document gives (what Copy and Export write). */
function documentSource(state: DocumentState, references: "eqref" | "plain"): string {
  return documentToMarkdown(state.doc, { formulaLatex: portableFormulaLatex, references, source: markdownSourceOf(state), sourceText: true });
}

function problemText(problem: SourceProblem, isEn: boolean): string {
  const what =
    problem.kind === "extra-brace"
      ? isEn ? "extra }" : "多余的 }"
      : problem.kind === "end"
        ? `${isEn ? "extra " : "多余的 "}\\end{${problem.detail}}`
        : problem.kind === "environment"
          ? `${isEn ? "unclosed " : "未闭合 "}\\begin{${problem.detail}}`
          : `${isEn ? "unclosed " : "未闭合 "}${problem.detail}`;
  return isEn ? `Line ${problem.line}: ${what}` : `第 ${problem.line} 行：${what}`;
}

/** The header's short list of what reading drops: names, then comments and preamble lines counted. */
function lossSummary(losses: SourceLoss[], isEn: boolean): string {
  const names = [...new Set(losses.filter((l) => l.kind === "command" || l.kind === "environment").map((l) => l.name))];
  const comments = losses.filter((l) => l.kind === "comment").length;
  const preamble = losses.filter((l) => l.kind === "preamble").length;
  const lines = losses.filter((l) => l.kind === "line").length;
  const parts = names.slice(0, 3);
  if (names.length > 3) parts.push("…");
  if (comments) parts.push(isEn ? `${comments} comment${comments > 1 ? "s" : ""}` : `注释 ${comments}`);
  if (preamble) parts.push(isEn ? `${preamble} preamble line${preamble > 1 ? "s" : ""}` : `导言 ${preamble} 行`);
  if (lines) parts.push(isEn ? `${lines} line${lines > 1 ? "s" : ""}` : `${lines} 行`);
  return (isEn ? (lines && !names.length ? "Rewritten: " : "Not kept: ") : lines && !names.length ? "将改写：" : "不保留：") + parts.join(isEn ? ", " : "、");
}

const lossTitle = (losses: SourceLoss[], isEn: boolean) =>
  losses.map((l) => `${isEn ? "Line" : "第"} ${l.lines.join(", ")}${isEn ? "" : " 行"}: ${l.name}`).join("\n");

export function DocumentSourceView(props: { theme: Theme }) {
  const hostRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorView | null>(null);
  const undoHistory = useRef(new Compartment()).current;
  const state = useDocumentEditorStore((store) => store.state);
  const references = useDocumentFormatStore((store) => store.markdownReferences);
  const setReferences = useDocumentFormatStore((store) => store.setMarkdownReferences);
  const uiLanguage = useEditorStore((store) => store.language);
  const isEn = uiLanguage === "en";
  // The source follows the document a moment after typing stops.
  const [doc, setDoc] = useState(state?.doc ?? null);
  useEffect(() => {
    const timer = window.setTimeout(() => setDoc(state?.doc ?? null), 250);
    return () => window.clearTimeout(timer);
  }, [state?.doc]);
  // The Markdown file the document was read from (its unchanged blocks are written back as they were).
  const markdownSource = state ? markdownSourceOf(state) : null;
  const source = useMemo(() => {
    if (!doc) return "";
    return documentToMarkdown(doc, { formulaLatex: portableFormulaLatex, references, source: markdownSource, sourceText: true });
  }, [doc, references, markdownSource]);

  // Being edited: the source editor has the focus (or the window went away while it had it). Its text is then the
  // user's, not the document's.
  const [editing, setEditing] = useState(false);
  // unfinished code that was not read: the text stays as typed until it is finished or put back
  const [problem, setProblem] = useState<SourceProblem | null>(null);
  const [losses, setLosses] = useState<SourceLoss[]>([]);
  const dirtyRef = useRef(false);
  const timerRef = useRef<number | null>(null);
  const editingRef = useRef(false);
  const referencesRef = useRef(references);
  referencesRef.current = references;

  const mark = useCallback((next: SourceProblem | null, lost: SourceLoss[]) => {
    viewRef.current?.dispatch({ effects: setMarks.of({ problem: next, lost: lost.flatMap((l) => l.lines) }) });
  }, []);

  /** The text put in from the document: only what differs replaced (the caret and scroll stay); its own undo cleared. */
  const showSource = useCallback(
    (text: string) => {
      const view = viewRef.current;
      if (!view) return;
      const current = view.state.doc.toString();
      if (current !== text) {
        let start = 0;
        const max = Math.min(current.length, text.length);
        while (start < max && current.charCodeAt(start) === text.charCodeAt(start)) start++;
        let end = 0;
        while (end < max - start && current.charCodeAt(current.length - 1 - end) === text.charCodeAt(text.length - 1 - end)) end++;
        view.dispatch({
          changes: { from: start, to: current.length - end, insert: text.slice(start, text.length - end) },
          annotations: [fromDocument.of(true), Transaction.addToHistory.of(false)],
        });
      }
      // the source editor's undo is for what is typed in it: the document's text starts it afresh
      view.dispatch({ effects: undoHistory.reconfigure([]) });
      view.dispatch({ effects: undoHistory.reconfigure(history()) });
      mark(null, []);
    },
    [mark, undoHistory],
  );

  /** What was typed, read into the document. False: unfinished, not read. */
  const apply = useCallback(
    (): boolean => {
      if (timerRef.current !== null) {
        window.clearTimeout(timerRef.current);
        timerRef.current = null;
      }
      const view = viewRef.current;
      const documentView = useDocumentEditorStore.getState().view;
      if (!view || !documentView || !dirtyRef.current) return true;
      const text = view.state.doc.toString();
      const unfinished = sourceProblem(text, "markdown");
      if (unfinished) {
        setProblem(unfinished);
        setLosses([]);
        mark(unfinished, []);
        return false;
      }
      const written = documentSource(documentView.state, referencesRef.current);
      try {
        const tr = sourceTransaction(documentView.state, text, "markdown", written);
        if (tr) documentView.dispatch(tr);
      } catch (error) {
        console.error("source not read", error);
        return false;
      }
      dirtyRef.current = false;
      const now = documentSource(documentView.state, referencesRef.current);
      const lost = markdownLosses(text, now);
      setProblem(null);
      setLosses(lost);
      mark(null, lost);
      return true;
    },
    [mark],
  );

  /** The editor left: what was typed read in; the document's source shown again (unless it could not be read). */
  const finishEditing = useCallback(() => {
    if (!editingRef.current) return;
    editingRef.current = false;
    apply();
    const documentState = useDocumentEditorStore.getState().state;
    setDoc(documentState?.doc ?? null);
    setEditing(false);
  }, [apply]);

  useEffect(() => {
    if (!hostRef.current) return;
    const view = new EditorView({
      parent: hostRef.current,
      state: EditorState.create({
        doc: source,
        extensions: [
          lineNumbers(),
          highlightActiveLineGutter(),
          undoHistory.of(history()),
          // a new line keeps the line's indentation (the language's own would indent a section's text)
          keymap.of([{ key: "Enter", run: insertNewlineKeepIndent }, ...defaultKeymap, ...historyKeymap]),
          syntaxHighlighting(highlightStyle),
          marksField,
          EditorView.lineWrapping,
          EditorView.contentAttributes.of({ spellcheck: "false" }),
          EditorView.updateListener.of((update) => {
            if (update.focusChanged) {
              if (update.view.hasFocus) {
                editingRef.current = true;
                setEditing(true);
              } else if (document.hasFocus()) finishEditing();
              // (the window went away: the text is read, and stays as typed for when it comes back)
              else apply();
            }
            if (!update.docChanged || update.transactions.every((tr) => tr.annotation(fromDocument))) return;
            dirtyRef.current = true;
            editingRef.current = true;
            setEditing(true);
            if (timerRef.current !== null) window.clearTimeout(timerRef.current);
            timerRef.current = window.setTimeout(() => {
              timerRef.current = null;
              apply();
            }, APPLY_DELAY_MS);
          }),
          theme,
        ],
      }),
    });
    viewRef.current = view;
    return () => {
      if (timerRef.current !== null) window.clearTimeout(timerRef.current);
      view.destroy();
      viewRef.current = null;
    };
    // The view is made once; text and language follow below.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // The focus moved on after the window came back elsewhere than the source editor: the editing is over.
  useEffect(() => {
    if (!editing) return;
    const onFocusIn = (event: FocusEvent) => {
      const view = viewRef.current;
      if (view && event.target instanceof Node && !view.dom.contains(event.target)) finishEditing();
    };
    document.addEventListener("focusin", onFocusIn);
    return () => document.removeEventListener("focusin", onFocusIn);
  }, [editing, finishEditing]);

  // the document's source, unless the user's text is being edited or holds unfinished code
  useEffect(() => {
    if (editing || problem) return;
    showSource(source);
    setLosses([]);
  }, [source, editing, problem, showSource]);

  // A click or the caret elsewhere in the document: the source shows that part (its caret stays).
  const selectionBlock = state ? state.selection.$from.index(0) : -1;
  useEffect(() => {
    const view = viewRef.current;
    if (!view || editing || problem || !doc || selectionBlock < 0 || view.state.doc.toString() !== source) return;
    const at = sourcePlaceOf(doc, selectionBlock, source);
    if (at !== null) view.dispatch({ effects: EditorView.scrollIntoView(at, { y: "nearest", yMargin: 24 }) });
  }, [selectionBlock, source, doc, editing, problem]);

  const revert = () => {
    const documentState = useDocumentEditorStore.getState().state;
    dirtyRef.current = false;
    editingRef.current = false;
    setEditing(false);
    setProblem(null);
    setLosses([]);
    if (documentState) showSource(documentSource(documentState, references));
  };

  /** The caret to the first line listed (the header's list of what is dropped, the unfinished piece). */
  const showLine = (line: number) => {
    const view = viewRef.current;
    if (!view || line < 1 || line > view.state.doc.lines) return;
    const at = view.state.doc.line(line).from;
    view.dispatch({ selection: { anchor: at }, effects: EditorView.scrollIntoView(at, { y: "center" }) });
    view.focus();
  };

  return (
    <div
      className="latex-source-editor document-source-view"
      data-theme={props.theme}
      data-editing={editing ? "true" : undefined}
      aria-label={isEn ? "Document source" : "文档源码"}
    >
      <div className="document-source-view-bar">
        <span className="document-source-view-status">
          {problem ? (
            <>
              <button
                type="button"
                className="document-source-view-note is-problem"
                title={isEn ? "Not applied" : "未应用"}
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => showLine(problem.line)}
              >
                {problemText(problem, isEn)}
              </button>
              <button type="button" className="document-source-view-action" onClick={revert} title={isEn ? "Back to the document's source" : "恢复为文档的源码"}>
                {isEn ? "Revert" : "还原"}
              </button>
            </>
          ) : losses.length ? (
            <button
              type="button"
              className="document-source-view-note is-lost"
              title={lossTitle(losses, isEn)}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => showLine(losses[0].lines[0] ?? 1)}
            >
              {lossSummary(losses, isEn)}
            </button>
          ) : null}
        </span>
        <span className="document-format-switch" role="radiogroup" aria-label={isEn ? "References to formulas" : "公式引用"}>
          <span className="document-format-switch-label">{isEn ? "References" : "公式引用"}</span>
          {(["eqref", "plain"] as const).map((value) => (
            <button
              key={value}
              type="button"
              role="radio"
              aria-checked={references === value}
              className={references === value ? "is-active" : ""}
              disabled={!!problem}
              title={
                value === "eqref"
                  ? "\\eqref + \\label"
                  : isEn
                    ? "Plain text: (1)"
                    : "纯文本：(1)"
              }
              onClick={() => setReferences(value)}
            >
              {value === "eqref" ? "\\eqref" : "(n)"}
            </button>
          ))}
        </span>
      </div>
      <div ref={hostRef} className="document-source-view-host" />
    </div>
  );
}
