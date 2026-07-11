import { useEffect, useRef, useState } from "react";
import { EditorState } from "@codemirror/state";
import { EditorView, keymap, lineNumbers } from "@codemirror/view";
import { defaultKeymap, history, historyKeymap } from "@codemirror/commands";
import { latex as latexLanguageSupport } from "codemirror-lang-latex";
import { Check, Code2, Copy, RotateCcw } from "lucide-react";
import { useEditorStore } from "../stores/editorStore";
import type { LatexCodeFormat } from "../types/formula";

interface Props {
  latex: string;
  theme: "light" | "dark";
  format: LatexCodeFormat;
  onApply: (latex: string, sourceFormat: LatexCodeFormat) => void;
  onCopy: () => void;
}

export function LatexSourceEditor({
  latex,
  theme,
  format,
  onApply,
  onCopy,
}: Props) {
  const hostRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorView | null>(null);
  const draftRef = useRef(latex);
  const sourceRef = useRef(latex);
  const dirtyRef = useRef(false);
  const suppressChangeRef = useRef(false);
  const formatRef = useRef(format);
  const onApplyRef = useRef(onApply);
  const formatRefreshFrameRef = useRef<number | null>(null);
  const [dirty, setDirty] = useState(false);
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  onApplyRef.current = onApply;

  const updateDirty = (value: boolean) => {
    dirtyRef.current = value;
    setDirty(value);
  };

  useEffect(() => {
    sourceRef.current = latex;
  }, [latex]);

  useEffect(() => {
    if (!hostRef.current) return;

    const editorTheme = EditorView.theme({
      "&": { backgroundColor: "transparent", color: "var(--text)" },
      ".cm-content": { caretColor: "var(--accent)", fontFamily: "'SFMono-Regular', Menlo, Consolas, monospace", fontSize: "12px", padding: "10px 0" },
      ".cm-gutters": { backgroundColor: "transparent", color: "var(--text-faint)", border: "none" },
      ".cm-activeLine": { backgroundColor: "color-mix(in srgb, var(--accent-soft) 38%, transparent)" },
      ".cm-focused": { outline: "none" },
      ".cm-selectionBackground, ::selection": { backgroundColor: "color-mix(in srgb, var(--accent) 22%, transparent) !important" },
    });

    const state = EditorState.create({
      doc: sourceRef.current,
      extensions: [
        lineNumbers(),
        history(),
        latexLanguageSupport({ enableLinting: false, enableTooltips: false }),
        keymap.of([...defaultKeymap, ...historyKeymap]),
        editorTheme,
        EditorView.lineWrapping,
        EditorView.updateListener.of((update) => {
          if (!update.docChanged) return;
          draftRef.current = update.state.doc.toString();
          if (suppressChangeRef.current) {
            suppressChangeRef.current = false;
            updateDirty(false);
            return;
          }
          updateDirty(draftRef.current !== sourceRef.current);
        }),
      ],
    });

    const view = new EditorView({ state, parent: hostRef.current });
    viewRef.current = view;
    draftRef.current = sourceRef.current;
    updateDirty(false);

    return () => {
      if (formatRefreshFrameRef.current !== null) {
        window.cancelAnimationFrame(formatRefreshFrameRef.current);
        formatRefreshFrameRef.current = null;
      }
      view.destroy();
      viewRef.current = null;
    };
  }, [theme]);

  useEffect(() => {
    const view = viewRef.current;
    if (!view || dirtyRef.current) return;
    const current = view.state.doc.toString();
    if (current === latex) return;

    suppressChangeRef.current = true;
    view.dispatch({
      changes: { from: 0, to: current.length, insert: latex },
    });
    draftRef.current = latex;
  }, [latex]);

  useEffect(() => {
    const previousFormat = formatRef.current;
    if (previousFormat === format) return;
    formatRef.current = format;

    if (!dirtyRef.current) return;
    onApplyRef.current(draftRef.current, previousFormat);
    updateDirty(false);

    if (formatRefreshFrameRef.current !== null) {
      window.cancelAnimationFrame(formatRefreshFrameRef.current);
    }
    formatRefreshFrameRef.current = window.requestAnimationFrame(() => {
      formatRefreshFrameRef.current = window.requestAnimationFrame(() => {
        const view = viewRef.current;
        if (!view) return;
        const nextSource = sourceRef.current;
        const current = view.state.doc.toString();
        if (current !== nextSource) {
          suppressChangeRef.current = true;
          view.dispatch({
            changes: { from: 0, to: current.length, insert: nextSource },
          });
        }
        draftRef.current = nextSource;
        updateDirty(false);
        formatRefreshFrameRef.current = null;
      });
    });
  }, [format]);

  const replaceDraft = (value: string) => {
    const view = viewRef.current;
    if (!view) return;
    const current = view.state.doc.toString();
    suppressChangeRef.current = true;
    view.dispatch({ changes: { from: 0, to: current.length, insert: value } });
    draftRef.current = value;
    updateDirty(false);
  };

  const applyDraft = () => {
    onApply(draftRef.current, formatRef.current);
    sourceRef.current = draftRef.current;
    updateDirty(false);
  };

  return (
    <section className="source-panel">
      <div className="source-panel-header">
        <div className="source-title">
          <Code2 size={16} />
          <span>{isEn ? "LaTeX source" : "LaTeX 源码"}</span>
          {dirty && (
            <span className="unsaved-chip">
              {isEn ? "Unsynced changes" : "有未同步更改"}
            </span>
          )}
        </div>
        <div className="source-actions">
          {dirty && (
            <>
              <button type="button" className="text-button" onClick={() => replaceDraft(latex)}>
                <RotateCcw size={14} /> {isEn ? "Reset" : "还原"}
              </button>
              <button type="button" className="primary-small-button" onClick={applyDraft}>
                <Check size={14} /> {isEn ? "Apply" : "同步到公式"}
              </button>
            </>
          )}
          <button
            type="button"
            className="text-button source-copy-button"
            onClick={onCopy}
            aria-label={isEn ? "Copy LaTeX source" : "复制 LaTeX 源码"}
            title={isEn ? "Copy LaTeX source" : "复制 LaTeX 源码"}
          >
            <Copy size={14} />
          </button>
        </div>
      </div>
      <div ref={hostRef} className="codemirror-host" />
    </section>
  );
}
