/// <reference path="../types/mathliveKernel.d.ts" />
import {
  acceptCompletion,
  autocompletion,
  snippet,
  type Completion,
  type CompletionContext,
  type CompletionResult,
} from "@codemirror/autocomplete";
import { Prec, type Extension } from "@codemirror/state";
import { EditorView, keymap } from "@codemirror/view";
import { visualTexLatexCommandSignatures } from "mathlive";
import { commandRegistry } from "../autocomplete/commandRegistry";
import { COMMAND_FREQUENCY } from "../autocomplete/commandFrequency.generated";
import { latestSharedCommands } from "../autocomplete/latestSharedCommands";
import { convertVisualTexLatexToMarkup } from "../editor/mathLiveIntegralCompatibility";
import { inspectMathLiveSourceSafety } from "../editor/mathLiveSourceSafety";
import { nativeSuggestionPreviewLatex } from "../editor/nativeSuggestionPreviews";

// Source-code completion for the LaTeX panel. Same commands and ranking as
// the visual editor's candidate panel (Overleaf's approach: CodeMirror's
// fuzzy match tier, then corpus frequency as boost), but inserted as
// CodeMirror snippets: every argument is an empty {} field the caret visits
// with Tab, so \frac gives \frac{|}{} instead of a bare command.

interface SourceCompletion extends Completion {
  preview: string;
}

const COMMAND_PATTERN = /\\(?:begin\{[A-Za-z*]*|[A-Za-z]*)$/;
const VALID_FOR = /^\\(?:begin\{[A-Za-z*]*|[A-Za-z]*)$/;

// Environments the catalog does not provide as templates.
const EXTRA_ENVIRONMENTS: Record<string, string> = {
  aligned: "\\begin{aligned}#{} & #{}\\\\ #{} & #{}\\end{aligned}",
  gathered: "\\begin{gathered}#{}\\\\ #{}\\end{gathered}",
  array: "\\begin{array}{#{cc}}#{} & #{}\\end{array}",
  smallmatrix: "\\begin{smallmatrix}#{} & #{}\\\\ #{} & #{}\\end{smallmatrix}",
};

/**
 * VisualTeX template -> CodeMirror snippet. \placeholder{} becomes a field;
 * LaTeX's own \{ and \} keep their backslash through the snippet parser,
 * which unescapes \{ to {.
 */
function snippetTemplate(template: string) {
  return template
    .replace(/\\placeholder\{\}/g, "\u0000")
    .replace(/\\([{}])/g, "\\\\$1")
    .replace(/\u0000/g, "#{}");
}

function nativeTemplate(name: string, required: number) {
  return name + "{#{}}".repeat(required);
}

function nativePreview(name: string, required: number) {
  const registered = nativeSuggestionPreviewLatex(name);
  if (registered) return registered;
  const samples = ["a", "b", "c", "d"];
  return name + samples.slice(0, required).map((sample) => `{${sample}}`).join("");
}

let cachedOptions: SourceCompletion[] | null = null;

function completionOptions(): SourceCompletion[] {
  if (cachedOptions) return cachedOptions;
  const options: SourceCompletion[] = [];
  const seen = new Set<string>();
  const variants = new Set(latestSharedCommands.map((command) => command.id));
  const frequency = (command: string) => COMMAND_FREQUENCY[command] ?? 0;
  const push = (
    label: string,
    template: string,
    preview: string,
    boost: number,
  ) => {
    const key = label + "\u0001" + template;
    if (seen.has(key)) return;
    seen.add(key);
    options.push({
      label,
      type: "latex",
      boost,
      preview,
      apply: template.includes("#{") ? snippet(template) : template,
    });
  };

  // Commands with a VisualTeX template keep only that editable form, except
  // big operators, which are meaningful bare as well (same as the visual
  // editor).
  const templated = new Set<string>();
  for (const command of commandRegistry) {
    const name = command.command.trim();
    if (!/^\\(?:begin\{[A-Za-z*]+\}|[A-Za-z]+)$/.test(name)) continue;
    if (command.insertTemplate.includes("\\placeholder{}")) templated.add(name);
    const tieBreak = variants.has(command.id) ? -0.002 : 0.002;
    push(
      name,
      snippetTemplate(command.insertTemplate),
      command.previewLatex,
      frequency(name) + tieBreak + command.defaultPriority * 1e-6,
    );
  }

  const bareOperators = new Set(["\\int", "\\iint", "\\iiint", "\\oint", "\\oiint", "\\oiiint", "\\sum", "\\prod", "\\lim"]);
  for (const [name, required] of visualTexLatexCommandSignatures()) {
    if (!/^\\[A-Za-z]+$/.test(name)) continue;
    if (templated.has(name) && !bareOperators.has(name)) continue;
    push(name, nativeTemplate(name, required), nativePreview(name, required), frequency(name));
  }

  for (const [environment, template] of Object.entries(EXTRA_ENVIRONMENTS)) {
    const label = `\\begin{${environment}}`;
    push(label, template, template.replace(/#\{([^}]*)\}/g, (_, text) => text || "x"), frequency(label));
  }

  cachedOptions = options;
  return options;
}

function latexSourceCompletions(context: CompletionContext): CompletionResult | null {
  const word = context.matchBefore(COMMAND_PATTERN);
  if (!word) return null;
  // The backslash of a row break (\\) does not start a command.
  let slashes = 0;
  for (let at = word.from - 1; at >= 0 && context.state.sliceDoc(at, at + 1) === "\\"; at--) slashes++;
  if (slashes % 2 === 1) return null;
  if (word.text.length < 2 && !context.explicit) return null;
  return { from: word.from, options: completionOptions(), validFor: VALID_FOR };
}

const previewMarkup = new Map<string, string>();

function renderPreview(completion: Completion) {
  const element = document.createElement("span");
  element.className = "cm-vt-completion-preview";
  const latex = (completion as SourceCompletion).preview;
  if (!latex) return element;
  let markup = previewMarkup.get(latex);
  if (markup === undefined) {
    markup = "";
    if (!inspectMathLiveSourceSafety(latex)) {
      try {
        // Text style keeps limits and fractions inside one fixed-height row.
        markup = convertVisualTexLatexToMarkup(`{\\textstyle ${latex}}`, { defaultMode: "math" });
      } catch {
        markup = "";
      }
    }
    previewMarkup.set(latex, markup);
  }
  element.innerHTML = markup;
  return element;
}

export const latexSourceCompletion: Extension = [
  autocompletion({
    override: [latexSourceCompletions],
    activateOnTyping: true,
    icons: false,
    maxRenderedOptions: 60,
    tooltipClass: () => "cm-vt-completion",
    optionClass: () => "cm-vt-completion-option",
    addToOptions: [{ render: renderPreview, position: 60 }],
  }),
  // Tab also accepts while the list is open; otherwise it moves between
  // snippet fields or indents as before.
  Prec.highest(keymap.of([{ key: "Tab", run: acceptCompletion }])),
];

export const latexSourceCompletionTheme = EditorView.theme({
  // The tooltip element is only an invisible positioning frame: when space
  // is tight CodeMirror pins its height and keeps it after the list gets
  // shorter, which would show as empty space. The visible box is the list,
  // which is always exactly as tall as its rows and sits against the caret.
  ".cm-tooltip.cm-tooltip-autocomplete.cm-vt-completion": {
    display: "flex",
    flexDirection: "column",
    minWidth: "260px",
    maxWidth: "min(420px, calc(100vw - 24px))",
    border: "none",
    backgroundColor: "transparent",
    boxShadow: "none",
    overflow: "visible",
    padding: "0",
    pointerEvents: "none",
  },
  ".cm-tooltip.cm-tooltip-above.cm-tooltip-autocomplete.cm-vt-completion": {
    justifyContent: "flex-end",
  },
  // CodeMirror anchors the panel at the glyph box, which ends inside the
  // 1.62 line height: keep it clear of the line being typed.
  ".cm-tooltip.cm-tooltip-below.cm-vt-completion": { marginTop: "6px" },
  ".cm-tooltip.cm-tooltip-above.cm-vt-completion": { marginTop: "-6px" },
  ".cm-tooltip.cm-tooltip-autocomplete.cm-vt-completion > ul": {
    flex: "0 1 auto",
    height: "auto",
    minHeight: "0",
    maxHeight: "calc(8 * 34px + 10px)",
    boxSizing: "border-box",
    border: "1px solid var(--border-strong)",
    borderRadius: "8px",
    backgroundColor: "var(--surface)",
    boxShadow: "var(--shadow-md)",
    padding: "4px",
    pointerEvents: "auto",
    fontFamily: "'SFMono-Regular', Menlo, Consolas, monospace",
  },
  ".cm-tooltip-autocomplete.cm-vt-completion > ul > li.cm-vt-completion-option": {
    height: "34px",
    display: "flex",
    alignItems: "center",
    gap: "12px",
    padding: "0 8px",
    borderRadius: "5px",
    color: "var(--text)",
    lineHeight: "34px",
  },
  ".cm-tooltip-autocomplete.cm-vt-completion > ul > li[aria-selected]": {
    backgroundColor: "var(--accent-soft)",
    color: "var(--text)",
  },
  ".cm-vt-completion .cm-completionLabel": {
    flex: "1 1 auto",
    minWidth: "0",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    fontSize: "12.5px",
  },
  ".cm-vt-completion .cm-completionMatchedText": {
    textDecoration: "none",
    color: "var(--accent)",
    fontWeight: "650",
  },
  ".cm-vt-completion .cm-vt-completion-preview": {
    flex: "0 0 auto",
    maxWidth: "190px",
    height: "30px",
    display: "flex",
    alignItems: "center",
    overflow: "hidden",
    fontSize: "15px",
    lineHeight: "1",
  },
});
