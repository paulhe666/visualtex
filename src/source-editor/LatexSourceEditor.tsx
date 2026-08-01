import { useEffect, useRef, useState } from "react";
import { EditorState } from "@codemirror/state";
import { HighlightStyle, syntaxHighlighting } from "@codemirror/language";
import { EditorView, keymap, lineNumbers } from "@codemirror/view";
import { tags } from "@lezer/highlight";
import { defaultKeymap, history, historyKeymap } from "@codemirror/commands";
import { latex as latexLanguageSupport } from "codemirror-lang-latex";
import {
  AlertTriangle,
  Code2,
  Copy,
  PanelBottomClose,
  RotateCcw,
} from "lucide-react";
import { useEditorStore } from "../stores/editorStore";
import type { LatexSourceDraftResult } from "../clipboard/LatexCopyService";
import type { LatexCodeFormat, Theme } from "../types/formula";

const visualTeXLatexHighlightStyle = HighlightStyle.define([
  {
    tag: [
      tags.keyword,
      tags.definitionKeyword,
      tags.macroName,
      tags.labelName,
      tags.heading,
    ],
    color: "var(--syntax-command)",
  },
  {
    tag: [tags.className, tags.typeName, tags.namespace],
    color: "var(--syntax-function)",
  },
  {
    tag: [tags.operator, tags.processingInstruction],
    color: "var(--syntax-operator)",
  },
  { tag: tags.number, color: "var(--syntax-number)" },
  { tag: tags.bracket, color: "var(--syntax-bracket)" },
  { tag: [tags.string, tags.quote, tags.meta], color: "var(--syntax-string)" },
  { tag: tags.comment, color: "var(--syntax-comment)", fontStyle: "italic" },
  { tag: [tags.variableName, tags.content], color: "var(--syntax-variable)" },
  { tag: tags.invalid, color: "var(--syntax-error)", textDecoration: "underline" },
]);

interface Props {
  latex: string;
  theme: Theme;
  format: LatexCodeFormat;
  onCollapse: () => void;
  showCollapseAction?: boolean;
  showCopyAction?: boolean;
  compact?: boolean;
  onLiveChange: (
    latex: string,
    sourceFormat: LatexCodeFormat,
  ) => LatexSourceDraftResult;
  onCopy: () => void;
}

export function LatexSourceEditor({
  latex,
  theme,
  format,
  onCollapse,
  showCollapseAction = true,
  showCopyAction = true,
  compact = false,
  onLiveChange,
  onCopy,
}: Props) {
  const hostRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorView | null>(null);
  const draftRef = useRef(latex);
  const sourceRef = useRef(latex);
  const latestLatexRef = useRef(latex);
  const dirtyRef = useRef(false);
  const sourceFocusedRef = useRef(false);
  const syncErrorRef = useRef<string | null>(null);
  const suppressChangeRef = useRef(false);
  const formatRef = useRef(format);
  const onLiveChangeRef = useRef(onLiveChange);
  const formatRefreshFrameRef = useRef<number | null>(null);
  const [dirty, setDirty] = useState(false);
  const [syncError, setSyncError] = useState<string | null>(null);
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  onLiveChangeRef.current = onLiveChange;

  const updateDirty = (value: boolean) => {
    dirtyRef.current = value;
    setDirty(value);
  };

  const updateSyncError = (value: string | null) => {
    syncErrorRef.current = value;
    setSyncError(value);
  };

  useEffect(() => {
    sourceRef.current = latex;
    latestLatexRef.current = latex;
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
        syntaxHighlighting(visualTeXLatexHighlightStyle),
        keymap.of([...defaultKeymap, ...historyKeymap]),
        editorTheme,
        EditorView.lineWrapping,
        EditorView.updateListener.of((update) => {
          if (update.focusChanged) {
            sourceFocusedRef.current = update.view.hasFocus;
            if (!update.view.hasFocus && !syncErrorRef.current) {
              window.requestAnimationFrame(() => {
                const view = viewRef.current;
                if (!view || view.hasFocus || syncErrorRef.current) return;
                const canonical = latestLatexRef.current;
                const current = view.state.doc.toString();
                if (current !== canonical) {
                  suppressChangeRef.current = true;
                  view.dispatch({
                    changes: { from: 0, to: current.length, insert: canonical },
                  });
                }
                draftRef.current = canonical;
                updateDirty(false);
              });
            }
          }
          if (!update.docChanged) return;
          draftRef.current = update.state.doc.toString();
          if (suppressChangeRef.current) {
            suppressChangeRef.current = false;
            updateDirty(false);
            updateSyncError(null);
            return;
          }
          const result = onLiveChangeRef.current(
            draftRef.current,
            formatRef.current,
          );
          updateSyncError(result.valid ? null : result.error ?? "invalid-latex");
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
    if (!view || sourceFocusedRef.current || syncErrorRef.current) return;
    const current = view.state.doc.toString();
    if (current === latex) {
      updateDirty(false);
      return;
    }

    suppressChangeRef.current = true;
    view.dispatch({
      changes: { from: 0, to: current.length, insert: latex },
    });
    draftRef.current = latex;
    updateDirty(false);
  }, [latex]);

  useEffect(() => {
    const previousFormat = formatRef.current;
    if (previousFormat === format) return;
    formatRef.current = format;
    updateSyncError(null);

    if (formatRefreshFrameRef.current !== null) {
      window.cancelAnimationFrame(formatRefreshFrameRef.current);
    }
    formatRefreshFrameRef.current = window.requestAnimationFrame(() => {
      formatRefreshFrameRef.current = window.requestAnimationFrame(() => {
        const view = viewRef.current;
        if (!view) return;
        const nextSource = latestLatexRef.current;
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
    const result = onLiveChangeRef.current(value, formatRef.current);
    updateSyncError(result.valid ? null : result.error ?? "invalid-latex");
    updateDirty(false);
  };

  const showHeader = !compact || dirty || Boolean(syncError);

  return (
    <section
      className={
        "source-panel" +
        (compact ? " is-compact" : "") +
        (compact && (dirty || syncError) ? " has-dirty-actions" : "") +
        (syncError ? " has-source-error" : "")
      }
    >
      {showHeader && (
        <div className="source-panel-header">
          {!compact && (
            <div className="source-title">
              <Code2 size={16} />
              <span>{isEn ? "LaTeX source" : "LaTeX 源码"}</span>
              {syncError ? (
                <span className="source-error-chip">
                  <AlertTriangle size={12} />
                  {syncError === "incomplete-format-wrapper"
                    ? isEn
                      ? "Formula wrapper is incomplete"
                      : "公式环境包裹尚未完成"
                    : isEn
                      ? "Incomplete fragment is shown as LaTeX"
                      : "未完成片段按源码显示，其余保持渲染"}
                </span>
              ) : dirty ? (
                <span className="source-live-chip">
                  {isEn ? "Live synced" : "已实时同步"}
                </span>
              ) : null}
            </div>
          )}
          {compact && syncError && (
            <span className="source-error-chip source-error-chip-compact">
              <AlertTriangle size={12} />
              {syncError === "incomplete-format-wrapper"
                ? isEn
                  ? "Incomplete wrapper"
                  : "环境包裹未完成"
                : isEn
                  ? "Partial LaTeX preview"
                  : "局部源码预览"}
            </span>
          )}
          <div className="source-actions">
            {(dirty || syncError) && (
              <button
                type="button"
                className="text-button"
                onClick={() => replaceDraft(latex)}
              >
                <RotateCcw size={14} /> {isEn ? "Reset" : "还原"}
              </button>
            )}
            {showCopyAction && (
              <button
                type="button"
                className="text-button source-copy-button"
                onClick={onCopy}
                aria-label={isEn ? "Copy LaTeX source" : "复制 LaTeX 源码"}
                title={isEn ? "Copy LaTeX source" : "复制 LaTeX 源码"}
              >
                <Copy size={14} />
              </button>
            )}
            {showCollapseAction && (
              <button
                type="button"
                className="text-button source-collapse-button"
                onClick={onCollapse}
                aria-label={isEn ? "Hide LaTeX source" : "收起 LaTeX 源码"}
                title={isEn ? "Hide LaTeX source" : "收起 LaTeX 源码"}
              >
                <PanelBottomClose size={14} />
              </button>
            )}
          </div>
        </div>
      )}
      <div ref={hostRef} className="codemirror-host" />
    </section>
  );
}
