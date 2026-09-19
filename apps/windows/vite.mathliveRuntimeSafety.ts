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
 * Apply the narrow runtime corrections VisualTeX needs while it remains on
 * MathLive 0.109.2:
 *
 * 1. Match MathLive 0.110.0's text/MathML escaping fix for CVE-2026-54705.
 * 2. Avoid writing `mode` through an absent root child when options are changed
 *    on an empty, already-mounted mathfield.
 * 3. Replace the whole model for replaceAll instead of range-deleting atoms.
 *    Range deletion could skip empty structures and retain a root environment.
 * 4. Preserve upright bold uppercase Greek when serializing nested atoms.
 * 5. Keep matrix pointer hits in the selected row when atom bounds overlap.
 * 6. Include sentinel-only structures in selection, deletion and serialization.
 *
 * Every replacement is guarded by an exact single-match assertion. A future
 * MathLive source change therefore fails the build instead of silently shipping
 * an incomplete compatibility patch.
 */
export function patchVisualTexMathLiveRuntimeSafety(source: string) {
  let patched = source;

  // An empty structural atom still occupies a selectable model position.
  // Its sentinel-only branch must not exclude it from getAtoms(), otherwise
  // select-all/delete, replacement and selection serialization silently skip it.
  patched = replaceExactly(
    patched,
    [
      '  const firstChild = includeFirstAtoms ? atom.firstChild : firstNonFirstChild(atom);',
      '  if (!firstChild) return false;',
    ].join("\n"),
    [
      '  const firstChild = includeFirstAtoms ? atom.firstChild : firstNonFirstChild(atom);',
      '  if (!firstChild) return true;',
    ].join("\n"),
    "empty structural atom selection",
  );

  // Array hit testing first selects a row, then erroneously visits children
  // from every row again. Overlapping accent bounds in a later row can win a
  // zero-distance tie over the fraction denominator that was actually clicked.
  // Only matrix-level scripts remain outside the row traversal.
  patched = replaceExactly(
    patched,
    [
      '    if (!atom.isMultiline) {',
      '      for (const child of atom.children) {',
      '        const r = nearestAtomFromPointRecursive(mathfield, cache, child, x, y);',
      '        if (r[0] <= result[0]) result = r;',
      '      }',
      '    }',
    ].join("\n"),
    [
      '    if (!atom.isMultiline) {',
      '      for (const branch of ["superscript", "subscript"]) {',
      '        for (const child of atom.branch(branch) || []) {',
      '          const r = nearestAtomFromPointRecursive(mathfield, cache, child, x, y);',
      '          if (r[0] <= result[0]) result = r;',
      '        }',
      '      }',
      '    }',
    ].join("\n"),
    "matrix row pointer hit testing",
  );

  // The serializer only recognizes ASCII alphanumerics as a mathbf run. An
  // accent must serialize its body again, so upright bold Delta becomes bm,
  // which loses the explicit upright intent used by subsequent style toggles.
  patched = replaceExactly(
    patched,
    '    if (/^[a-zA-Z0-9]+$/.test(value))',
    '    if (/^[a-zA-Z0-9]+$/.test(value) || (x.every((atom) => atom.style.variant === "normal" && atom.style.variantStyle === "bold") && /^[a-zA-Z0-9ΓΔΘΛΞΠΣΥΦΨΩ]+$/.test(value)))',
    "upright bold uppercase Greek serialization",
  );

  // getAtoms([0, -1]) previously excluded sentinel-only structures (e.g.
  // \\vec{}), leaving them behind to accumulate on macro reparse and setValue.
  // Even with selection corrected above, a document replacement must start
  // from a fresh root, including when the previous root was a multiline array.
  patched = replaceExactly(
    patched,
    '    else if (options.insertionMode === "replaceAll") model.deleteAtoms();',
    [
      '    else if (options.insertionMode === "replaceAll") {',
      '      model.root = new Atom({ type: "root", mode: "math", body: [] });',
      '      model.position = 0;',
      '    }',
    ].join("\n"),
    "complete model replacement",
  );

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
