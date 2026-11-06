// The source tab in document mode: the document as LaTeX or as Markdown — the output format, chosen in this tab's
// header — read-only, following the document as it is edited. What it shows is what Copy and Export write.
import { useEffect, useMemo, useRef, useState } from "react";
import { Compartment, EditorState, type Extension } from "@codemirror/state";
import { EditorView, highlightActiveLineGutter, lineNumbers } from "@codemirror/view";
import { HighlightStyle, syntaxHighlighting } from "@codemirror/language";
import { tags } from "@lezer/highlight";
import { latex as latexLanguageSupport } from "codemirror-lang-latex";
import { documentToLatex, documentToMarkdown, markdownSourceOf } from "../../shared/visualtex-document/src/index.ts";
import { useDocumentEditorStore } from "../stores/documentEditorStore";
import { useDocumentFormatStore, type DocumentFormat } from "../stores/documentFormatStore";
import { useEditorStore } from "../stores/editorStore";
import type { Theme } from "../types/formula";
import { portableFormulaLatex } from "./DocumentEditor";

const highlightStyle = HighlightStyle.define([
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

const theme: Extension = EditorView.theme({
  "&": { backgroundColor: "transparent", color: "var(--text)", height: "100%" },
  ".cm-scroller": { overflow: "auto" },
  ".cm-content": {
    fontFamily: "'SFMono-Regular', Menlo, Consolas, monospace",
    fontSize: "var(--visualtex-latex-source-font-size, 12px)",
    lineHeight: "1.62",
    padding: "10px 0 18px",
  },
  ".cm-line": { paddingInline: "4px 10px" },
  ".cm-gutters": { backgroundColor: "transparent", color: "var(--text-faint)", border: "none" },
  ".cm-lineNumbers .cm-gutterElement": { minWidth: "32px", paddingInline: "6px 7px" },
  ".cm-activeLineGutter": { backgroundColor: "transparent" },
  ".cm-focused": { outline: "none" },
  ".cm-selectionBackground, ::selection": { backgroundColor: "color-mix(in srgb, var(--accent) 22%, transparent) !important" },
});

const languageOf = (format: DocumentFormat): Extension =>
  format === "latex" ? latexLanguageSupport({ enableLinting: false, enableTooltips: false, enableAutocomplete: false }) : [];

export function DocumentSourceView(props: { theme: Theme }) {
  const hostRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorView | null>(null);
  const language = useRef(new Compartment()).current;
  const state = useDocumentEditorStore((store) => store.state);
  const format = useDocumentFormatStore((store) => store.outputFormat);
  const setFormat = useDocumentFormatStore((store) => store.setOutputFormat);
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
    return format === "markdown"
      ? documentToMarkdown(doc, { formulaLatex: portableFormulaLatex, references, source: markdownSource })
      : documentToLatex(doc, { formulaLatex: portableFormulaLatex });
  }, [doc, format, references, markdownSource]);

  useEffect(() => {
    if (!hostRef.current) return;
    const view = new EditorView({
      parent: hostRef.current,
      state: EditorState.create({
        doc: source,
        extensions: [
          lineNumbers(),
          highlightActiveLineGutter(),
          language.of(languageOf(format)),
          syntaxHighlighting(highlightStyle),
          EditorState.readOnly.of(true),
          EditorView.editable.of(false),
          EditorView.lineWrapping,
          theme,
        ],
      }),
    });
    viewRef.current = view;
    return () => {
      view.destroy();
      viewRef.current = null;
    };
    // The view is made once; text and language follow below.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    viewRef.current?.dispatch({ effects: language.reconfigure(languageOf(format)) });
  }, [format, language]);

  useEffect(() => {
    const view = viewRef.current;
    if (!view || view.state.doc.toString() === source) return;
    view.dispatch({ changes: { from: 0, to: view.state.doc.length, insert: source } });
  }, [source]);

  return (
    <div
      className="latex-source-editor document-source-view"
      data-theme={props.theme}
      aria-label={isEn ? "Document source (read-only)" : "文档源码（只读）"}
    >
      <div className="document-source-view-bar">
        <span className="document-source-view-note">{isEn ? "Read-only" : "只读"}</span>
        {format === "markdown" && (
          <span className="document-format-switch" role="radiogroup" aria-label={isEn ? "References to formulas" : "公式引用"}>
            <span className="document-format-switch-label">{isEn ? "References" : "公式引用"}</span>
            {(["eqref", "plain"] as const).map((value) => (
              <button
                key={value}
                type="button"
                role="radio"
                aria-checked={references === value}
                className={references === value ? "is-active" : ""}
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
        )}
        <span className="document-format-switch" role="radiogroup" aria-label={isEn ? "Output format" : "输出格式"}>
          <span className="document-format-switch-label">{isEn ? "Output" : "输出"}</span>
          {(["latex", "markdown"] as const).map((value) => (
            <button
              key={value}
              type="button"
              role="radio"
              aria-checked={format === value}
              className={format === value ? "is-active" : ""}
              onClick={() => setFormat(value)}
            >
              {value === "latex" ? "LaTeX" : "Markdown"}
            </button>
          ))}
        </span>
      </div>
      <div ref={hostRef} className="document-source-view-host" />
    </div>
  );
}
