import { EditorState, Prec, type Extension } from "@codemirror/state";
import { insertNewlineAndIndent } from "@codemirror/commands";
import {
  indentOnInput,
  indentService,
  indentUnit,
  syntaxTree,
} from "@codemirror/language";
import {
  Decoration,
  EditorView,
  keymap,
  ViewPlugin,
  type DecorationSet,
} from "@codemirror/view";
import { latexLanguage } from "codemirror-lang-latex";
import { commandRegistry } from "../autocomplete/commandRegistry";
import type { CommandCategory } from "../types/command";

export const LATEX_SOURCE_INDENT = "  ";

interface LatexStructureState {
  depth: number;
  dollarFenceOpen: boolean;
}

function isEscaped(source: string, index: number) {
  let slashCount = 0;
  for (let cursor = index - 1; cursor >= 0 && source[cursor] === "\\"; cursor -= 1) {
    slashCount += 1;
  }
  return slashCount % 2 === 1;
}

function sourceBeforeComment(line: string) {
  for (let index = 0; index < line.length; index += 1) {
    if (line[index] === "%" && !isEscaped(line, index)) {
      return line.slice(0, index);
    }
  }
  return line;
}

function leadingClosingDepth(code: string, state: LatexStructureState) {
  let rest = code.trimStart();
  let closing = 0;

  if (state.dollarFenceOpen && rest === "$$") {
    return 1;
  }
  if (rest.startsWith("\\]")) {
    closing += 1;
    rest = rest.slice(2).trimStart();
  }

  while (true) {
    const environment = rest.match(/^\\end\{[A-Za-z]+\*?\}/);
    if (!environment) break;
    closing += 1;
    rest = rest.slice(environment[0].length).trimStart();
  }

  while (rest.startsWith("}")) {
    closing += 1;
    rest = rest.slice(1).trimStart();
  }

  return closing;
}

function scanStructureCode(code: string, state: LatexStructureState) {
  const trimmed = code.trim();
  if (!trimmed) return;

  if (trimmed === "$$") {
    if (state.dollarFenceOpen) {
      state.depth = Math.max(0, state.depth - 1);
      state.dollarFenceOpen = false;
    } else {
      state.depth += 1;
      state.dollarFenceOpen = true;
    }
    return;
  }
  if (trimmed === "\\[") {
    state.depth += 1;
    return;
  }
  if (trimmed === "\\]") {
    state.depth = Math.max(0, state.depth - 1);
    return;
  }

  let delta = 0;
  const environmentPattern = /\\(begin|end)\{[A-Za-z]+\*?\}/g;
  for (const match of trimmed.matchAll(environmentPattern)) {
    const index = match.index ?? 0;
    if (isEscaped(trimmed, index)) continue;
    delta += match[1] === "begin" ? 1 : -1;
  }

  for (let index = 0; index < trimmed.length; index += 1) {
    if (isEscaped(trimmed, index)) continue;
    if (trimmed[index] === "{") delta += 1;
    else if (trimmed[index] === "}") delta -= 1;
  }

  state.depth = Math.max(0, state.depth + delta);
}

function structureStateForSource(source: string) {
  const state: LatexStructureState = { depth: 0, dollarFenceOpen: false };
  for (const rawLine of source.replace(/\r\n?/g, "\n").split("\n")) {
    scanStructureCode(sourceBeforeComment(rawLine), state);
  }
  return state;
}

export function formatLatexSourceForEditor(source: string) {
  const normalized = String(source ?? "").replace(/\r\n?/g, "\n");
  const state: LatexStructureState = { depth: 0, dollarFenceOpen: false };
  return normalized
    .split("\n")
    .map((rawLine) => {
      if (!rawLine.trim()) return "";
      const content = rawLine.trimStart();
      const code = sourceBeforeComment(content);
      const indentDepth = Math.max(
        0,
        state.depth - leadingClosingDepth(code, state),
      );
      const formatted = LATEX_SOURCE_INDENT.repeat(indentDepth) + content;
      scanStructureCode(code, state);
      return formatted;
    })
    .join("\n");
}

function latexIndentationAt(
  state: EditorState,
  pos: number,
  contextLine: { text: string; from: number },
  indentWidth: number,
) {
  const prefix = state.doc.sliceString(0, contextLine.from);
  const structureState = structureStateForSource(prefix);
  const code = sourceBeforeComment(contextLine.text);
  const indentDepth = Math.max(
    0,
    structureState.depth - leadingClosingDepth(code, structureState),
  );
  return indentDepth * indentWidth;
}

const commandCategoryByControlSequence = new Map<string, CommandCategory>();
for (const command of commandRegistry) {
  const match = command.command.trim().match(/^\\([A-Za-z@]+)$/);
  if (!match || commandCategoryByControlSequence.has(match[1])) continue;
  commandCategoryByControlSequence.set(match[1], command.category);
}
commandCategoryByControlSequence.set("begin", "structure");
commandCategoryByControlSequence.set("end", "structure");

function semanticCommandClass(source: string) {
  const name = source.match(/^\\([A-Za-z@]+)/)?.[1];
  if (!name) return "cm-vt-command-default";
  const category = commandCategoryByControlSequence.get(name);
  return category
    ? `cm-vt-command-${category}`
    : "cm-vt-command-default";
}

function buildSourceDecorations(view: EditorView): DecorationSet {
  const ranges = [];

  for (const visible of view.visibleRanges) {
    let line = view.state.doc.lineAt(visible.from);
    while (line.from <= visible.to) {
      const whitespace = line.text.match(/^[ \t]+/)?.[0] ?? "";
      if (whitespace.length > 0) {
        ranges.push(
          Decoration.mark({ class: "cm-vt-indent-guide" }).range(
            line.from,
            line.from + whitespace.length,
          ),
        );
      }
      if (line.to >= visible.to || line.number >= view.state.doc.lines) break;
      line = view.state.doc.line(line.number + 1);
    }

    syntaxTree(view.state).iterate({
      from: visible.from,
      to: visible.to,
      enter(node) {
        const isControlSequence =
          node.name === "CtrlSeq" ||
          node.name === "Begin" ||
          node.name === "End" ||
          node.name.endsWith("CtrlSeq");
        if (!isControlSequence || node.from >= node.to) return;
        const source = view.state.doc.sliceString(node.from, node.to);
        ranges.push(
          Decoration.mark({ class: semanticCommandClass(source) }).range(
            node.from,
            node.to,
          ),
        );
      },
    });
  }

  ranges.sort((left, right) => left.from - right.from || left.to - right.to);
  return Decoration.set(ranges, true);
}

const sourceDecorations = ViewPlugin.fromClass(
  class {
    decorations: DecorationSet;

    constructor(view: EditorView) {
      this.decorations = buildSourceDecorations(view);
    }

    update(update: { docChanged: boolean; viewportChanged: boolean; view: EditorView }) {
      if (update.docChanged || update.viewportChanged) {
        this.decorations = buildSourceDecorations(update.view);
      }
    }
  },
  {
    decorations: (plugin) => plugin.decorations,
  },
);

function hasMatchingEnvironmentEndAfter(
  source: string,
  from: number,
  environmentName: string,
) {
  const trailing = source.slice(from);
  const tokenPattern = /\\(begin|end)\{([^{}\r\n]+)\}/g;
  let depth = 1;

  for (const match of trailing.matchAll(tokenPattern)) {
    const index = match.index ?? 0;
    if (isEscaped(trailing, index)) continue;
    const lineStart = trailing.lastIndexOf("\n", index - 1) + 1;
    const beforeToken = trailing.slice(lineStart, index);
    if (sourceBeforeComment(beforeToken).length !== beforeToken.length) continue;
    if (match[2].trim() !== environmentName) continue;

    depth += match[1] === "begin" ? 1 : -1;
    if (depth === 0) return true;
  }

  return false;
}

function insertNewlineWithoutDuplicateEnvironmentEnd(view: EditorView) {
  const selection = view.state.selection.main;
  if (!selection.empty) return false;

  const line = view.state.doc.lineAt(selection.head);
  const prefix = line.text.slice(0, selection.head - line.from);
  const suffix = line.text.slice(selection.head - line.from);
  if (suffix.trim()) return false;

  const begin = prefix.match(/\\begin\{([^{}\r\n]+)\}[ \t]*$/);
  if (!begin) return false;
  const environmentName = begin[1].trim();
  if (!environmentName) return false;

  const source = view.state.doc.toString();
  if (
    !hasMatchingEnvironmentEndAfter(
      source,
      selection.head,
      environmentName,
    )
  ) {
    return false;
  }

  return insertNewlineAndIndent(view);
}

const duplicateEnvironmentEndGuard = Prec.highest(
  keymap.of([
    { key: "Enter", run: insertNewlineWithoutDuplicateEnvironmentEnd },
  ]),
);

export const visualTeXLatexEditingExtensions: Extension = [
  EditorState.tabSize.of(2),
  indentUnit.of(LATEX_SOURCE_INDENT),
  indentService.of((context, pos) =>
    latexIndentationAt(context.state, pos, context.lineAt(pos, 1), context.unit),
  ),
  latexLanguage.data.of({
    indentOnInput: /^\s*(?:\\end\{[^}]*\}|\\\]|\$\$|\})/,
  }),
  indentOnInput(),
  duplicateEnvironmentEndGuard,
  sourceDecorations,
];
