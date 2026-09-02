const replacementCount = (source: string, target: string) =>
  source.split(target).length - 1;

function replaceExactly(
  source: string,
  target: string,
  replacement: string,
  label: string,
) {
  const count = replacementCount(source, target);
  if (count !== 1) {
    throw new Error(`MathLive ${label} patch anchor changed (${count}).`);
  }
  return source.replace(target, replacement);
}

export function escapeMathLiveTextForMarkup(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

export function escapeMathLiveTextForXml(value: string) {
  return escapeMathLiveTextForMarkup(value)
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Apply the two narrow runtime corrections VisualTeX needs while it remains on
 * MathLive 0.109.2:
 *
 * 1. Match MathLive 0.110.0's text/MathML escaping fix for CVE-2026-54705.
 * 2. Avoid writing `mode` through an absent root child when options are changed
 *    on an empty, already-mounted mathfield.
 *
 * Every replacement is guarded by an exact single-match assertion. A future
 * MathLive source change therefore fails the build instead of silently shipping
 * an incomplete compatibility patch.
 */
export function patchVisualTexMathLiveRuntimeSafety(source: string) {
  let patched = source;

  patched = replaceExactly(
    patched,
    '    let body = (_a3 = this.value) != null ? _a3 : "";',
    '    let body = this.value ? visualTexEscapeMathLiveText(this.value) : "";',
    "HTML text escaping",
  );
  patched = replaceExactly(
    patched,
    "function sanitizeAttributeName(attribute) {",
    [
      "function visualTexEscapeMathLiveText(value) {",
      '  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");',
      "}",
      "function sanitizeAttributeName(attribute) {",
    ].join("\n"),
    "HTML escape helper",
  );

  patched = replaceExactly(
    patched,
    [
      "function xmlEscape(string) {",
      `  return string.replace(/"/g, "&quot;").replace(/'/g, "&#39;").replace(/</g, "&lt;").replace(/>/g, "&gt;");`,
      "}",
    ].join("\n"),
    [
      "function xmlEscape(string) {",
      `  return string.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/'/g, "&#39;").replace(/</g, "&lt;").replace(/>/g, "&gt;");`,
      "}",
    ].join("\n"),
    "MathML XML escaping",
  );
  patched = replaceExactly(
    patched,
    [
      "    mathML = `<mtext ${makeID(",
      "      stream.atoms[initial].id,",
      "      options",
      "    )}>${mathML}</mtext>`;",
    ].join("\n"),
    [
      "    mathML = `<mtext ${makeID(",
      "      stream.atoms[initial].id,",
      "      options",
      "    )}>${xmlEscape(mathML)}</mtext>`;",
    ].join("\n"),
    "MathML scanned text escaping",
  );
  patched = replaceExactly(
    patched,
    [
      '  if (atom.mode === "text")',
      "    return `<mi${makeID(atom.id, options)}>${atom.value}</mi>`;",
    ].join("\n"),
    [
      '  if (atom.mode === "text")',
      "    return `<mi${makeID(atom.id, options)}>${xmlEscape(atom.value)}</mi>`;",
    ].join("\n"),
    "MathML text atom escaping",
  );

  const rawDelimiterReplacements: Array<[string, string, string]> = [
    [
      '          result += `<mo>${SPECIAL_DELIMS[arrayAtom.leftDelim] || arrayAtom.leftDelim}</mo>`;',
      '          result += `<mo>${xmlEscape(SPECIAL_DELIMS[arrayAtom.leftDelim] || arrayAtom.leftDelim)}</mo>`;',
      "MathML array left delimiter escaping",
    ],
    [
      '          result += `<mo>${SPECIAL_DELIMS[arrayAtom.rightDelim] || arrayAtom.rightDelim}</mo>`;',
      '          result += `<mo>${xmlEscape(SPECIAL_DELIMS[arrayAtom.rightDelim] || arrayAtom.rightDelim)}</mo>`;',
      "MathML array right delimiter escaping",
    ],
    [
      '        result += "<mo" + makeID(atom.id, options) + ">" + (SPECIAL_DELIMS[genfracAtom.leftDelim] || genfracAtom.leftDelim) + "</mo>";',
      [
        '        result += "<mo" + makeID(atom.id, options) + ">" + xmlEscape(',
        "          SPECIAL_DELIMS[genfracAtom.leftDelim] || genfracAtom.leftDelim",
        '        ) + "</mo>";',
      ].join("\n"),
      "MathML fraction left delimiter escaping",
    ],
    [
      '        result += "<mo" + makeID(atom.id, options) + ">" + (SPECIAL_DELIMS[genfracAtom.rightDelim] || genfracAtom.rightDelim) + "</mo>";',
      [
        '        result += "<mo" + makeID(atom.id, options) + ">" + xmlEscape(',
        "          SPECIAL_DELIMS[genfracAtom.rightDelim] || genfracAtom.rightDelim",
        '        ) + "</mo>";',
      ].join("\n"),
      "MathML fraction right delimiter escaping",
    ],
    [
      '        result += `<mo${makeID(atom.id, options)}>${(_a3 = SPECIAL_DELIMS[lDelim]) != null ? _a3 : lDelim}</mo>`;',
      [
        "        result += `<mo${makeID(atom.id, options)}>${xmlEscape(",
        "          (_a3 = SPECIAL_DELIMS[lDelim]) != null ? _a3 : lDelim",
        "        )}</mo>`;",
      ].join("\n"),
      "MathML left-right opening delimiter escaping",
    ],
    [
      '        result += `<mo${makeID(atom.id, options)}>${(_b3 = SPECIAL_DELIMS[rDelim]) != null ? _b3 : rDelim}</mo>`;',
      [
        "        result += `<mo${makeID(atom.id, options)}>${xmlEscape(",
        "          (_b3 = SPECIAL_DELIMS[rDelim]) != null ? _b3 : rDelim",
        "        )}</mo>`;",
      ].join("\n"),
      "MathML left-right closing delimiter escaping",
    ],
    [
      '      result += `<mo${makeID(atom.id, options)}>${SPECIAL_DELIMS[atom.value] || atom.value}</mo>`;',
      [
        "      result += `<mo${makeID(atom.id, options)}>${xmlEscape(",
        "        SPECIAL_DELIMS[atom.value] || atom.value",
        "      )}</mo>`;",
      ].join("\n"),
      "MathML sized delimiter escaping",
    ],
  ];
  for (const [target, replacement, label] of rawDelimiterReplacements) {
    patched = replaceExactly(patched, target, replacement, label);
  }

  patched = replaceExactly(
    patched,
    [
      '        if (typeof atom.value === "string" && atom.value.charCodeAt(0) > 255) {',
      '          result = "&#x" + ("000000" + atom.value.charCodeAt(0).toString(16)).slice(-4) + ";";',
      '        } else if (typeof atom.value === "string")',
    ].join("\n"),
    [
      '        if (typeof atom.value === "string" && atom.value.charCodeAt(0) > 255) {',
      "          result = String.fromCodePoint(atom.value.codePointAt(0));",
      '        } else if (typeof atom.value === "string")',
    ].join("\n"),
    "MathML Unicode character serialization",
  );
  patched = replaceExactly(
    patched,
    [
      '          if (typeof codepoint === "number") {',
      '            result = "&#x" + ("000000" + codepoint.toString(16)).slice(-4) + ";";',
      "          }",
    ].join("\n"),
    [
      '          if (typeof codepoint === "number") {',
      "            result = String.fromCodePoint(codepoint);",
      "          }",
    ].join("\n"),
    "MathML char command serialization",
  );
  patched = replaceExactly(
    patched,
    '      result += "<mtext" + makeID(atom.id, options) + ">" + atom.value + "</mtext>";',
    '      result += "<mtext" + makeID(atom.id, options) + ">" + xmlEscape(atom.value) + "</mtext>";',
    "MathML LaTeX atom escaping",
  );
  patched = replaceExactly(
    patched,
    '      result += `<mtext ${makeID(atom.id, options)}x>${atom.value}</mtext>`;',
    [
      "      result += `<mtext ${makeID(atom.id, options)}x>${xmlEscape(",
      "        atom.value",
      "      )}</mtext>`;",
    ].join("\n"),
    "MathML explicit text atom escaping",
  );

  patched = replaceExactly(
    patched,
    [
      '    if (((_a3 = this.model.root.firstChild) == null ? void 0 : _a3.mode) !== mode)',
      "      this.model.root.firstChild.mode = mode;",
    ].join("\n"),
    [
      "    const visualTexRootFirstChild = this.model.root.firstChild;",
      "    if (visualTexRootFirstChild && visualTexRootFirstChild.mode !== mode)",
      "      visualTexRootFirstChild.mode = mode;",
    ].join("\n"),
    "empty-model option mutation",
  );

  return patched;
}
