#!/usr/bin/env -S npx tsx

import assert from "node:assert/strict";
import { EditorState } from "@codemirror/state";
import { getIndentation, indentRange } from "@codemirror/language";
import { latex as latexLanguageSupport } from "codemirror-lang-latex";
import {
  formatLatexSourceForEditor,
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

console.log("LaTeX source editor formatting and indentation regression passed.");
