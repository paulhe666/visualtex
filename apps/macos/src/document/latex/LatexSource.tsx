// The source of a LaTeX document: one CodeMirror editor showing the file of the active tab (each file keeps its
// own state: text, undo, selection), with the look of the document source view. TeX's errors and warnings are
// marked on their lines; a line found from the pages flashes.
import { forwardRef, useEffect, useImperativeHandle, useRef } from "react";
import { EditorSelection, EditorState, StateEffect, StateField, type Extension } from "@codemirror/state";
import { Decoration, EditorView, highlightActiveLine, highlightActiveLineGutter, keymap, lineNumbers, type DecorationSet } from "@codemirror/view";
import { syntaxHighlighting } from "@codemirror/language";
import { defaultKeymap, history, historyKeymap, indentWithTab, insertNewlineKeepIndent } from "@codemirror/commands";
import { latex as latexLanguageSupport } from "codemirror-lang-latex";
import { highlightStyle, theme } from "../DocumentSourceView";
import { latexSourceCompletion, latexSourceCompletionTheme } from "../../source-editor/latexSourceCompletion";

export interface LineMark {
  line: number;
  kind: "error" | "warning" | "box";
  message: string;
}

export interface LatexSourceHandle {
  /** show a file (its saved state, or a new one with `text`; `readOnly`: not editable) */
  show(path: string, text: string, readOnly?: boolean): void;
  /** a file's current text (null: never opened) */
  text(path: string): string | null;
  /** replace a file's text from outside (a reload); undo kept */
  setText(path: string, text: string): void;
  active(): string | null;
  /** caret to a line of the active file, scrolled to the middle, the line flashing */
  goto(line: number): void;
  mark(marks: LineMark[]): void;
  caret(): { path: string; pos: number; line: number } | null;
  replace(from: number, to: number, insert: string): void;
  coords(pos: number): { left: number; top: number; bottom: number } | null;
  focus(): void;
  forget(path: string): void;
  /** caret to the end of a line of the active file (tests) */
  caretToLine(line: number, column?: number): void;
}

interface Props {
  onChange(path: string, text: string): void;
  onCaret(path: string, line: number, pos: number): void;
  /** Mod+E: the formula at the caret into the formula editor */
  onFormula(): void;
  /** Mod+S */
  onSave(): void;
}

const setMarks = StateEffect.define<LineMark[]>();
const flashLine = StateEffect.define<number | null>();

const marksField = StateField.define<DecorationSet>({
  create: () => Decoration.none,
  update(marks, tr) {
    marks = marks.map(tr.changes);
    for (const e of tr.effects) {
      if (!e.is(setMarks)) continue;
      const doc = tr.state.doc;
      const byLine = new Map<number, LineMark>();
      for (const m of e.value) {
        if (m.line < 1 || m.line > doc.lines) continue;
        const old = byLine.get(m.line);
        if (!old || (old.kind !== "error" && m.kind === "error")) byLine.set(m.line, m);
      }
      const ranges = [...byLine.values()]
        .sort((a, b) => a.line - b.line)
        .map((m) =>
          Decoration.line({
            class: m.kind === "error" ? "cm-vtx-problem-line cm-vtx-tex-error" : "cm-vtx-lost-line cm-vtx-tex-warning",
            attributes: { title: m.message },
          }).range(doc.line(m.line).from),
        );
      marks = Decoration.set(ranges, true);
    }
    return marks;
  },
  provide: (f) => EditorView.decorations.from(f),
});

const flashField = StateField.define<DecorationSet>({
  create: () => Decoration.none,
  update(deco, tr) {
    deco = deco.map(tr.changes);
    for (const e of tr.effects) {
      if (!e.is(flashLine)) continue;
      deco = e.value && e.value <= tr.state.doc.lines ? Decoration.set([Decoration.line({ class: "cm-vtx-tex-flash" }).range(tr.state.doc.line(e.value).from)]) : Decoration.none;
    }
    return deco;
  },
  provide: (f) => EditorView.decorations.from(f),
});

const texLanguage = latexLanguageSupport({ enableLinting: false, enableTooltips: false, enableAutocomplete: false, autoCloseTags: false });

export const LatexSource = forwardRef<LatexSourceHandle, Props>(function LatexSource(props, ref) {
  const hostRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorView | null>(null);
  const states = useRef(new Map<string, EditorState>());
  const activeRef = useRef<string | null>(null);
  const propsRef = useRef(props);
  propsRef.current = props;
  const flashTimer = useRef(0);

  const extensions = (path: string): Extension => [
    lineNumbers(),
    highlightActiveLineGutter(),
    highlightActiveLine(),
    history(),
    keymap.of([
      { key: "Mod-e", run: () => (propsRef.current.onFormula(), true) },
      { key: "Mod-s", run: () => (propsRef.current.onSave(), true), preventDefault: true },
      { key: "Enter", run: insertNewlineKeepIndent },
      indentWithTab,
      ...defaultKeymap,
      ...historyKeymap,
    ]),
    /\.(tex|ltx|sty|cls|def|cfg|clo|dtx)$/i.test(path) ? [texLanguage, latexSourceCompletion, latexSourceCompletionTheme] : [],
    syntaxHighlighting(highlightStyle),
    marksField,
    flashField,
    EditorView.lineWrapping,
    EditorView.contentAttributes.of({ spellcheck: "false", "aria-label": path }),
    EditorView.updateListener.of((update) => {
      const file = activeRef.current;
      if (!file || update.view !== viewRef.current) return;
      if (update.docChanged) propsRef.current.onChange(file, update.state.doc.toString());
      if (update.selectionSet || update.docChanged) {
        const head = update.state.selection.main.head;
        propsRef.current.onCaret(file, update.state.doc.lineAt(head).number, head);
      }
    }),
    theme,
  ];

  useEffect(() => {
    if (!hostRef.current) return;
    const view = new EditorView({ parent: hostRef.current, state: EditorState.create({ doc: "" }) });
    viewRef.current = view;
    return () => {
      view.destroy();
      viewRef.current = null;
    };
  }, []);

  useImperativeHandle(ref, () => ({
    show(path, text, readOnly = false) {
      const view = viewRef.current;
      if (!view) return;
      if (activeRef.current === path) return;
      if (activeRef.current) states.current.set(activeRef.current, view.state);
      let state = states.current.get(path);
      if (!state) {
        state = EditorState.create({ doc: text, extensions: [extensions(path), EditorState.readOnly.of(readOnly)] });
        states.current.set(path, state);
      }
      activeRef.current = path;
      view.setState(state);
    },
    text(path) {
      if (activeRef.current === path && viewRef.current) return viewRef.current.state.doc.toString();
      return states.current.get(path)?.doc.toString() ?? null;
    },
    setText(path, text) {
      const view = viewRef.current;
      if (activeRef.current === path && view) {
        const now = view.state.doc.toString();
        if (now !== text) view.dispatch({ changes: { from: 0, to: now.length, insert: text } });
        return;
      }
      const state = states.current.get(path);
      if (state && state.doc.toString() !== text) states.current.set(path, state.update({ changes: { from: 0, to: state.doc.length, insert: text } }).state);
    },
    active: () => activeRef.current,
    goto(line) {
      const view = viewRef.current;
      if (!view) return;
      const n = Math.min(Math.max(line, 1), view.state.doc.lines);
      const at = view.state.doc.line(n);
      view.dispatch({ selection: EditorSelection.cursor(at.from), effects: [EditorView.scrollIntoView(at.from, { y: "center" }), flashLine.of(n)] });
      window.clearTimeout(flashTimer.current);
      flashTimer.current = window.setTimeout(() => viewRef.current?.dispatch({ effects: flashLine.of(null) }), 1400);
    },
    mark(marks) {
      viewRef.current?.dispatch({ effects: setMarks.of(marks) });
    },
    caret() {
      const view = viewRef.current;
      const path = activeRef.current;
      if (!view || !path) return null;
      const pos = view.state.selection.main.head;
      return { path, pos, line: view.state.doc.lineAt(pos).number };
    },
    replace(from, to, insert) {
      const view = viewRef.current;
      if (!view) return;
      view.dispatch({ changes: { from, to, insert }, selection: EditorSelection.cursor(from + insert.length) });
    },
    coords(pos) {
      const c = viewRef.current?.coordsAtPos(pos);
      return c ? { left: c.left, top: c.top, bottom: c.bottom } : null;
    },
    focus: () => viewRef.current?.focus(),
    forget(path) {
      states.current.delete(path);
    },
    caretToLine(line, column) {
      const view = viewRef.current;
      if (!view) return;
      const l = view.state.doc.line(Math.min(Math.max(line, 1), view.state.doc.lines));
      const at = column === undefined ? l.to : Math.min(l.from + column, l.to);
      view.dispatch({ selection: EditorSelection.cursor(at), effects: EditorView.scrollIntoView(at, { y: "center" }) });
      view.focus();
    },
  }));

  return <div ref={hostRef} className="document-source-view-host vtx-tex-source-host" />;
});
