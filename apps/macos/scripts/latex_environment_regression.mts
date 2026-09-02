#!/usr/bin/env -S npx tsx

import assert from "node:assert/strict";
import {
  isSingleCompleteLatexEnvironment,
  unwrapSingleLatexDisplayMath,
} from "../src/math/latexEnvironment";

const accepted = [
  String.raw`\begin{aligned}a&=b\\c&=d\end{aligned}`,
  String.raw`
    \begin{aligned}
      a &= \begin{cases}b,&c\\d,&e\end{cases}\\
      f &= g
    \end{aligned}
  `,
  String.raw`% heading comment
\begin{aligned}
a&=b % \end{aligned} inside a comment
\\ c&=d
\end{aligned} % trailing comment`,
  String.raw`\begin{aligned}a&=100\%\\b&=c\end{aligned}`,
];

const rejected = [
  String.raw`a=b\\c=d`,
  String.raw`\begin{aligned}a&=b`,
  String.raw`\begin{aligned}a&=b\end{gathered}`,
  String.raw`\begin{aligned}a&=b\end{aligned}+c`,
  String.raw`\begin{aligned}a&=b\end{aligned}\begin{aligned}c&=d\end{aligned}`,
  String.raw`\\begin{aligned}a&=b\\end{aligned}`,
];

for (const source of accepted) {
  assert.equal(
    isSingleCompleteLatexEnvironment(source),
    true,
    `Expected one complete environment: ${source}`,
  );
}
for (const source of rejected) {
  assert.equal(
    isSingleCompleteLatexEnvironment(source),
    false,
    `Expected source to remain splittable/rejected: ${source}`,
  );
}

const wrappedAligned = String.raw`\[
\begin{aligned}
a&=b\\c&=d
\end{aligned}
\]`;
assert.equal(
  unwrapSingleLatexDisplayMath(wrappedAligned),
  String.raw`\begin{aligned}
a&=b\\c&=d
\end{aligned}`,
);
assert.equal(
  unwrapSingleLatexDisplayMath(String.raw`$$\begin{align*}a&=b\end{align*}$$`),
  String.raw`\begin{align*}a&=b\end{align*}`,
);
assert.equal(unwrapSingleLatexDisplayMath(String.raw`$$a$$b$$`), null);
assert.equal(unwrapSingleLatexDisplayMath(String.raw`\(a=b\)`), null);

console.log("LaTeX complete-environment regression PASS");
