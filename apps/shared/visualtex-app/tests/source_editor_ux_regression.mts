#!/usr/bin/env -S npx tsx

import assert from "node:assert/strict";
import { EditorSelection, EditorState, type Transaction } from "@codemirror/state";
import { getIndentation, indentRange } from "@codemirror/language";
import { latex as latexLanguageSupport } from "codemirror-lang-latex";
import {
  environmentNameCompletionEnd,
  formatLatexSourceForEditor,
  insertNewlineClosingEnvironment,
  visualTeXLatexEditingExtensions,
} from "../src/source-editor/latexSourceEditorSupport.ts";

const unindented = [
  String.raw`\begin{align}`,
  String.raw`\frac{\alpha+1}{\beta}&=\int_0^1 x^2\,\mathrm{d}x \\`,
  String.raw`\begin{matrix}`,
  String.raw`a&b \\`,
  String.raw`c&d`,
  String.raw`\end{matrix}&\approx\gamma`,
  String.raw`\end{align}`,
].join("\n");

const expected = [
  String.raw`\begin{align}`,
  String.raw`  \frac{\alpha+1}{\beta}&=\int_0^1 x^2\,\mathrm{d}x \\`,
  String.raw`  \begin{matrix}`,
  String.raw`    a&b \\`,
  String.raw`    c&d`,
  String.raw`  \end{matrix}&\approx\gamma`,
  String.raw`\end{align}`,
].join("\n");

assert.equal(
  formatLatexSourceForEditor(unindented),
  expected,
  "generated LaTeX source must use deterministic environment indentation",
);

const displayMath = [
  "$$",
  String.raw`\begin{cases}`,
  String.raw`x&x>0 \\`,
  String.raw`0&x\le 0`,
  String.raw`\end{cases}`,
  "$$",
].join("\n");
const expectedDisplayMath = [
  "$$",
  String.raw`  \begin{cases}`,
  String.raw`    x&x>0 \\`,
  String.raw`    0&x\le 0`,
  String.raw`  \end{cases}`,
  "$$",
].join("\n");
assert.equal(formatLatexSourceForEditor(displayMath), expectedDisplayMath);

const commentsDoNotIndent = [
  String.raw`% \begin{matrix}`,
  String.raw`\begin{equation}`,
  String.raw`x=1 % \begin{matrix}`,
  String.raw`\end{equation}`,
].join("\n");
assert.equal(
  formatLatexSourceForEditor(commentsDoNotIndent),
  [
    String.raw`% \begin{matrix}`,
    String.raw`\begin{equation}`,
    String.raw`  x=1 % \begin{matrix}`,
    String.raw`\end{equation}`,
  ].join("\n"),
);

const state = EditorState.create({
  doc: unindented,
  extensions: [
    latexLanguageSupport({ enableLinting: false, enableTooltips: false }),
    visualTeXLatexEditingExtensions,
  ],
});

const expectedIndentColumns = [0, 2, 2, 4, 4, 2, 0];
for (let lineNumber = 1; lineNumber <= state.doc.lines; lineNumber += 1) {
  const line = state.doc.line(lineNumber);
  assert.equal(
    getIndentation(state, line.from),
    expectedIndentColumns[lineNumber - 1],
    `unexpected indentation for line ${lineNumber}: ${line.text}`,
  );
}

const changes = indentRange(state, 0, state.doc.length);
const reindented = changes.apply(state.doc).toString();
assert.equal(
  reindented,
  expected,
  "CodeMirror indentRange must use the same LaTeX environment indentation as generated source",
);

const multilineGroup = [
  String.raw`\frac{`,
  "a+b",
  "}{",
  "c+d",
  "}",
].join("\n");
assert.equal(
  formatLatexSourceForEditor(multilineGroup),
  [
    String.raw`\frac{`,
    "  a+b",
    "}{",
    "  c+d",
    "}",
  ].join("\n"),
  "multiline command arguments should indent and dedent like code blocks",
);

// Enter right after \begin{name}: an indented line for the caret and \end{name} below, unless the \end is there
const pressEnter = (doc: string, caret = doc.length) => {
  let next: EditorState | null = null;
  const handled = insertNewlineClosingEnvironment({
    state: EditorState.create({
      doc,
      selection: EditorSelection.cursor(caret),
      extensions: [visualTeXLatexEditingExtensions],
    }),
    dispatch: (transaction: Transaction) => {
      next = transaction.state;
    },
  });
  const state = next as EditorState | null;
  return handled && state
    ? { doc: state.doc.toString(), caret: state.selection.main.head }
    : null;
};
assert.deepEqual(pressEnter(String.raw`\begin{align*}`), {
  doc: [String.raw`\begin{align*}`, "  ", String.raw`\end{align*}`].join("\n"),
  caret: String.raw`\begin{align*}`.length + 3,
});
assert.deepEqual(pressEnter(["$$", String.raw`  \begin{cases}`, "$$"].join("\n"), 18), {
  doc: ["$$", String.raw`  \begin{cases}`, "    ", String.raw`  \end{cases}`, "$$"].join("\n"),
  caret: 23,
});
const withEnd = [String.raw`\begin{matrix}`, String.raw`\end{matrix}`].join("\n");
assert.equal(
  pressEnter(withEnd, String.raw`\begin{matrix}`.length)?.doc,
  [String.raw`\begin{matrix}`, "  ", String.raw`\end{matrix}`].join("\n"),
  "an \\end that is already there is not added again",
);
assert.equal(pressEnter(String.raw`\begin{matrix} a`), null, "text after the \\begin: an ordinary new line");
assert.equal(pressEnter(String.raw`% \begin{matrix}`), null, "a commented \\begin is not an environment");
assert.equal(pressEnter(String.raw`a \\begin{x}`), null, "a row break before begin{x} is not \\begin");

// a completion of \begin{name| replaces the brace the editor paired with it, never a brace of an outer group
const completionEnd = (doc: string, from: number, pos: number) =>
  environmentNameCompletionEnd(EditorState.create({ doc }), from, pos);
assert.equal(completionEnd(String.raw`  \begin{matr}`, 2, 13), 14);
assert.equal(completionEnd(String.raw`\frac{\begin{matr}}`, 6, 17), 18);
assert.equal(completionEnd(String.raw`\frac{\begin{matr}`, 6, 17), 17, "the brace closes \\frac{");
assert.equal(completionEnd(String.raw`\begin{matr`, 0, 11), 11);
assert.equal(completionEnd(String.raw`\alpha}`, 0, 6), 6, "only an environment name takes its brace");

console.log("LaTeX source editor formatting and indentation regression passed.");
