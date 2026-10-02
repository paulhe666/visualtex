import { useEffect, useRef, useState, type CSSProperties } from "react";
import { Annotation, EditorState, Transaction } from "@codemirror/state";
import {
  foldGutter,
  foldKeymap,
  HighlightStyle,
  syntaxHighlighting,
} from "@codemirror/language";
import {
  EditorView,
  highlightActiveLine,
  highlightActiveLineGutter,
  highlightSpecialChars,
  keymap,
  lineNumbers,
} from "@codemirror/view";
import { tags } from "@lezer/highlight";
import {
  defaultKeymap,
  history,
  historyKeymap,
  indentWithTab,
} from "@codemirror/commands";
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
import type { LatexCodeFormat } from "../types/formula";
import { visualTeXLatexEditingExtensions } from "./latexSourceEditorSupport";

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
    fontWeight: "650",
  },
  {
    tag: [tags.strong, tags.emphasis, tags.monospace],
    color: "var(--syntax-command)",
    fontWeight: "600",
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

const externalSourceSync = Annotation.define<boolean>();
interface SourceDraftStatus {
  dirty: boolean;
  error: string | null;
  hasLivePreview: boolean;
}
const syncedDraft: SourceDraftStatus = { dirty: false, error: null, hasLivePreview: true };

interface Props {
  latex: string;
  format: LatexCodeFormat;
  onCollapse: () => void;
  showCollapseAction?: boolean;
  showCopyAction?: boolean;
  compact?: boolean;
  onLiveChange: (
    latex: string,
    sourceFormat: LatexCodeFormat,
  ) => LatexSourceDraftResult;
  onFocusChange?: (focused: boolean) => void;
  onCopy: () => void;
  forceExternalSync?: boolean;
}

export function LatexSourceEditor({
  latex,
  format,
  onCollapse,
  showCollapseAction = true,
  showCopyAction = true,
  compact = false,
  onLiveChange,
  onFocusChange,
  onCopy,
  forceExternalSync = false,
}: Props) {
  const hostRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorView | null>(null);
  // CodeMirror owns draft text. This baseline is the last accepted source;
  // canonicalSourceRef is the current serialization of the visual document.
  const acceptedSourceRef = useRef(latex);
  const canonicalSourceRef = useRef(latex);
  canonicalSourceRef.current = latex;
  const statusRef = useRef(syncedDraft);
  const [status, setStatus] = useState(syncedDraft);
  const { dirty, error: syncError, hasLivePreview } = status;
  const formatRef = useRef(format);
  const onLiveChangeRef = useRef(onLiveChange);
  const onFocusChangeRef = useRef(onFocusChange);
  const focusReleaseFrameRef = useRef<number | null>(null);
  const language = useEditorStore((state) => state.language);
  const sourceEditorFontSize = useEditorStore(
    (state) => state.sourceEditorFontSize,
  );
  const isEn = language === "en";
  onLiveChangeRef.current = onLiveChange;
  onFocusChangeRef.current = onFocusChange;

  const updateStatus = (next: SourceDraftStatus) => {
    statusRef.current = next;
    setStatus(next);
  };

  const publishDraft = (value: string) => {
    const result = onLiveChangeRef.current(value, formatRef.current);
    const preview = result.valid || result.values.length > 0 || Boolean(result.previewValues?.length);
    if (preview) acceptedSourceRef.current = value;
    updateStatus({
      dirty: !preview && value !== acceptedSourceRef.current,
      error: result.valid ? null : result.error ?? "invalid-latex",
      hasLivePreview: preview,
    });
  };

  const syncDocument = (view: EditorView, value: string) => {
    const current = view.state.doc.toString();
    if (current !== value) {
      view.dispatch({
        changes: { from: 0, to: current.length, insert: value },
        annotations: [externalSourceSync.of(true), Transaction.addToHistory.of(false)],
      });
    }
    acceptedSourceRef.current = value;
    updateStatus(syncedDraft);
  };

  useEffect(() => {
    if (!hostRef.current) return;

    const editorTheme = EditorView.theme({
      "&": {
        backgroundColor: "transparent",
        color: "var(--text)",
      },
      ".cm-content": {
        caretColor: "var(--accent)",
        fontFamily: "'SFMono-Regular', Menlo, Consolas, monospace",
        fontSize: "var(--source-editor-font-size, 12px)",
        lineHeight: "1.62",
        padding: "10px 0 18px",
      },
      ".cm-line": {
        paddingInline: "4px 10px",
      },
      ".cm-gutters": {
        backgroundColor: "transparent",
        color: "var(--text-faint)",
        border: "none",
      },
      ".cm-lineNumbers .cm-gutterElement": {
        minWidth: "32px",
        paddingInline: "6px 7px",
      },
      ".cm-foldGutter .cm-gutterElement": {
        width: "18px",
        paddingInline: "1px",
        color: "var(--text-faint)",
        cursor: "pointer",
      },
      ".cm-activeLineGutter": {
        backgroundColor: "color-mix(in srgb, var(--accent-soft) 45%, transparent)",
        color: "var(--text-muted)",
      },
      ".cm-activeLine": {
        backgroundColor: "color-mix(in srgb, var(--accent-soft) 32%, transparent)",
      },
      ".cm-focused": { outline: "none" },
      ".cm-selectionBackground, ::selection": {
        backgroundColor:
          "color-mix(in srgb, var(--accent) 22%, transparent) !important",
      },
      ".cm-matchingBracket": {
        backgroundColor: "color-mix(in srgb, var(--syntax-bracket) 18%, transparent)",
        outline: "1px solid color-mix(in srgb, var(--syntax-bracket) 55%, transparent)",
        borderRadius: "2px",
      },
      ".cm-vt-indent-guide": {
        backgroundImage:
          "linear-gradient(to right, transparent calc(2ch - 1px), color-mix(in srgb, var(--border-strong) 58%, transparent) calc(2ch - 1px), color-mix(in srgb, var(--border-strong) 58%, transparent) 2ch)",
        backgroundSize: "2ch 100%",
      },
      ".cm-vt-command-default, .cm-vt-command-structure, .cm-vt-command-calculus, .cm-vt-command-matrix, .cm-vt-command-greek, .cm-vt-command-relation, .cm-vt-command-set, .cm-vt-command-arrow, .cm-vt-command-physics": {
        fontWeight: "600",
      },
      ".cm-vt-command-default, .cm-vt-command-structure": {
        color: "var(--syntax-command)",
      },
      ".cm-vt-command-calculus": { color: "var(--syntax-number)" },
      ".cm-vt-command-matrix": { color: "var(--syntax-function)" },
      ".cm-vt-command-greek": { color: "var(--syntax-string)" },
      ".cm-vt-command-relation": { color: "var(--syntax-error)" },
      ".cm-vt-command-set": { color: "var(--syntax-bracket)" },
      ".cm-vt-command-arrow": {
        color:
          "color-mix(in srgb, var(--syntax-command) 45%, var(--syntax-function))",
      },
      ".cm-vt-command-physics": {
        color:
          "color-mix(in srgb, var(--syntax-bracket) 55%, var(--syntax-string))",
      },
      ".cm-tooltip-autocomplete": {
        border: "1px solid var(--border-strong)",
        borderRadius: "7px",
        backgroundColor: "var(--bg-elevated)",
        boxShadow: "var(--shadow-md)",
        overflow: "hidden",
      },
      ".cm-tooltip-autocomplete > ul > li[aria-selected]": {
        backgroundColor: "var(--accent-soft)",
        color: "var(--text)",
      },
    });

    const state = EditorState.create({
      doc: canonicalSourceRef.current,
      extensions: [
        lineNumbers(),
        foldGutter({ openText: "⌄", closedText: "›" }),
        highlightActiveLineGutter(),
        highlightActiveLine(),
        highlightSpecialChars(),
        history(),
        keymap.of([indentWithTab, ...foldKeymap, ...defaultKeymap, ...historyKeymap]),
        latexLanguageSupport({ enableLinting: false, enableTooltips: false }),
        visualTeXLatexEditingExtensions,
        syntaxHighlighting(visualTeXLatexHighlightStyle),
        EditorView.contentAttributes.of({
          spellcheck: "false",
          autocapitalize: "off",
          autocomplete: "off",
          "data-gramm": "false",
        }),
        editorTheme,
        EditorView.lineWrapping,
        EditorView.updateListener.of((update) => {
          if (update.focusChanged) {
            if (focusReleaseFrameRef.current !== null) {
              window.cancelAnimationFrame(focusReleaseFrameRef.current);
              focusReleaseFrameRef.current = null;
            }
            if (update.view.hasFocus) {
              onFocusChangeRef.current?.(true);
            } else {
              focusReleaseFrameRef.current = window.requestAnimationFrame(() => {
                focusReleaseFrameRef.current = null;
                const view = viewRef.current;
                if (!view || view.hasFocus) return;
                onFocusChangeRef.current?.(false);
                if (statusRef.current.error || !statusRef.current.hasLivePreview) return;
                syncDocument(view, canonicalSourceRef.current);
              });
            }
          }
          if (!update.docChanged || update.transactions.some(tr => tr.annotation(externalSourceSync))) return;
          publishDraft(update.state.doc.toString());
        }),
      ],
    });

    const view = new EditorView({ state, parent: hostRef.current });
    viewRef.current = view;

    return () => {
      if (focusReleaseFrameRef.current !== null) {
        window.cancelAnimationFrame(focusReleaseFrameRef.current);
        focusReleaseFrameRef.current = null;
      }
      if (view.hasFocus) {
        onFocusChangeRef.current?.(false);
      }
      view.destroy();
      viewRef.current = null;
    };
  }, []);

  useEffect(() => {
    const view = viewRef.current;
    const formatChanged = formatRef.current !== format;
    formatRef.current = format;
    if (!view) return;
    if (!forceExternalSync && !formatChanged && (view.hasFocus || statusRef.current.error)) return;
    syncDocument(view, latex);
  }, [forceExternalSync, format, latex]);

  const replaceDraft = (value: string) => {
    const view = viewRef.current;
    if (!view) return;
    syncDocument(view, value);
    publishDraft(value);
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
      style={
        {
          "--source-editor-font-size": `${sourceEditorFontSize}px`,
        } as CSSProperties
      }
      data-source-editor-font-size={sourceEditorFontSize}
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
                      ? "Preview only — source validation has not passed"
                      : "仅预览：源码尚未通过校验"}
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
                  ? "Unvalidated source preview"
                  : "源码未通过校验，仅预览"}
            </span>
          )}
          <div className="source-actions">
            {(dirty || (syncError && !hasLivePreview)) && (
              <button
                type="button"
                className="text-button"
                data-source-reset
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
