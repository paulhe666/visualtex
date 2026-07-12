import { useEffect, useRef } from "react";
import { basicSetup } from "codemirror";
import { autocompletion, type CompletionContext, type CompletionResult } from "@codemirror/autocomplete";
import { StreamLanguage } from "@codemirror/language";
import { stex } from "@codemirror/legacy-modes/mode/stex";
import { EditorState, Prec } from "@codemirror/state";
import { EditorView, keymap } from "@codemirror/view";

export interface SourceChange {
  startByte: number;
  endByte: number;
  replacement: string;
}

export interface SourceReveal {
  line: number;
  column?: number | null;
  requestId: number;
}

export type SourceCompletionKind = "label" | "citation" | "command";

export interface SourceCompletion {
  label: string;
  kind: SourceCompletionKind;
  detail?: string;
}

export interface SourceCompletionContext {
  kind: SourceCompletionKind;
  fromOffset: number;
}

export interface SourceEditorProps {
  value: string;
  readOnly?: boolean;
  reveal?: SourceReveal | null;
  completions?: SourceCompletion[];
  onChange: (change: SourceChange, nextValue: string) => void;
  onUndo?: () => void;
  onRedo?: () => void;
  onCursor?: (utf16Offset: number) => void;
}

const utf8Length = (value: string): number => new TextEncoder().encode(value).length;

export function completionContextAt(linePrefix: string): SourceCompletionContext | null {
  const reference = linePrefix.match(/\\(?:ref|eqref|autoref|cref|Cref)\{([^{}]*)$/);
  if (reference) {
    const value = reference[1] ?? "";
    const segment = value.slice(value.lastIndexOf(",") + 1);
    return { kind: "label", fromOffset: linePrefix.length - segment.length };
  }
  const citation = linePrefix.match(/\\(?:cite|citep|citet|parencite|textcite)\{([^{}]*)$/);
  if (citation) {
    const value = citation[1] ?? "";
    const segment = value.slice(value.lastIndexOf(",") + 1);
    return { kind: "citation", fromOffset: linePrefix.length - segment.length };
  }
  const command = linePrefix.match(/\\[A-Za-z@]*$/);
  if (command?.index !== undefined) {
    return { kind: "command", fromOffset: command.index };
  }
  return null;
}

const isLowSurrogate = (value: number): boolean => value >= 0xdc00 && value <= 0xdfff;
const isHighSurrogate = (value: number): boolean => value >= 0xd800 && value <= 0xdbff;

export function computeSourceChange(previous: string, next: string): SourceChange | null {
  if (previous === next) return null;

  let start = 0;
  const sharedLength = Math.min(previous.length, next.length);
  while (start < sharedLength && previous.charCodeAt(start) === next.charCodeAt(start)) {
    start += 1;
  }
  if (
    start > 0 &&
    (isLowSurrogate(previous.charCodeAt(start)) || isLowSurrogate(next.charCodeAt(start)))
  ) {
    start -= 1;
  }

  let previousEnd = previous.length;
  let nextEnd = next.length;
  while (
    previousEnd > start &&
    nextEnd > start &&
    previous.charCodeAt(previousEnd - 1) === next.charCodeAt(nextEnd - 1)
  ) {
    previousEnd -= 1;
    nextEnd -= 1;
  }
  if (
    previousEnd < previous.length &&
    previousEnd > start &&
    isHighSurrogate(previous.charCodeAt(previousEnd - 1))
  ) {
    previousEnd += 1;
    nextEnd += 1;
  }

  return {
    startByte: utf8Length(previous.slice(0, start)),
    endByte: utf8Length(previous.slice(0, previousEnd)),
    replacement: next.slice(start, nextEnd),
  };
}

export function SourceEditor({
  value,
  readOnly = false,
  reveal = null,
  completions = [],
  onChange,
  onUndo,
  onRedo,
  onCursor,
}: SourceEditorProps) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const viewRef = useRef<EditorView | null>(null);
  const valueRef = useRef(value);
  const suppressRef = useRef(false);
  const completionsRef = useRef(completions);
  completionsRef.current = completions;
  const callbacksRef = useRef({ onChange, onUndo, onRedo, onCursor });
  callbacksRef.current = { onChange, onUndo, onRedo, onCursor };

  useEffect(() => {
    if (!hostRef.current) return;
    const undoRedo = Prec.highest(
      keymap.of([
        {
          key: "Mod-z",
          run: () => {
            callbacksRef.current.onUndo?.();
            return Boolean(callbacksRef.current.onUndo);
          },
        },
        {
          key: "Shift-Mod-z",
          run: () => {
            callbacksRef.current.onRedo?.();
            return Boolean(callbacksRef.current.onRedo);
          },
        },
        {
          key: "Mod-y",
          run: () => {
            callbacksRef.current.onRedo?.();
            return Boolean(callbacksRef.current.onRedo);
          },
        },
      ]),
    );
    const completionSource = (context: CompletionContext): CompletionResult | null => {
      const line = context.state.doc.lineAt(context.pos);
      const prefix = context.state.sliceDoc(line.from, context.pos);
      const completionContext = completionContextAt(prefix);
      if (!completionContext) return null;
      const options = completionsRef.current
        .filter((completion) => completion.kind === completionContext.kind)
        .map((completion) => ({
          label: completion.label,
          apply: completion.label,
          detail: completion.detail,
          type:
            completion.kind === "command"
              ? "function"
              : completion.kind === "citation"
                ? "text"
                : "variable",
        }));
      if (options.length === 0 && !context.explicit) return null;
      return {
        from: line.from + completionContext.fromOffset,
        options,
        validFor: completionContext.kind === "command" ? /\\?[A-Za-z@]*/ : /[^{},]*/,
      };
    };
    const state = EditorState.create({
      doc: valueRef.current,
      extensions: [
        basicSetup,
        autocompletion({ override: [completionSource], activateOnTyping: true }),
        StreamLanguage.define(stex),
        EditorState.readOnly.of(readOnly),
        EditorView.lineWrapping,
        undoRedo,
        EditorView.updateListener.of((update) => {
          if (update.selectionSet) {
            callbacksRef.current.onCursor?.(update.state.selection.main.head);
          }
          if (!update.docChanged || suppressRef.current) return;
          const nextValue = update.state.doc.toString();
          const change = computeSourceChange(valueRef.current, nextValue);
          valueRef.current = nextValue;
          if (change) callbacksRef.current.onChange(change, nextValue);
        }),
        EditorView.theme({
          "&": { height: "100%", fontSize: "14px" },
          ".cm-scroller": { overflow: "auto", fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace" },
          ".cm-content": { padding: "16px 0 40px" },
          ".cm-gutters": { background: "var(--panel-subtle)", border: "none" },
          "&.cm-focused": { outline: "none" },
        }),
      ],
    });
    const view = new EditorView({ state, parent: hostRef.current });
    viewRef.current = view;
    return () => {
      view.destroy();
      viewRef.current = null;
    };
  }, [readOnly]);

  useEffect(() => {
    const view = viewRef.current;
    if (!view || view.state.doc.toString() === value) {
      valueRef.current = value;
      return;
    }
    suppressRef.current = true;
    view.dispatch({
      changes: { from: 0, to: view.state.doc.length, insert: value },
    });
    valueRef.current = value;
    suppressRef.current = false;
  }, [value]);

  useEffect(() => {
    const view = viewRef.current;
    if (!view || !reveal) return;
    const lineNumber = Math.max(1, Math.min(view.state.doc.lines, reveal.line));
    const line = view.state.doc.line(lineNumber);
    const requestedColumn = Math.max(1, reveal.column ?? 1);
    const position = Math.min(line.to, line.from + requestedColumn - 1);
    view.dispatch({
      selection: { anchor: position },
      effects: EditorView.scrollIntoView(position, { y: "center" }),
    });
    view.focus();
  }, [reveal?.requestId, reveal?.line, reveal?.column]);

  return <div ref={hostRef} className="vt-source-editor" aria-label="LaTeX source editor" />;
}
