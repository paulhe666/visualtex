/**
 * CP254 isolated parser candidate.
 *
 * Input must be the current owned MathLive kernel AFTER the complete production
 * Vite transformation. This file is not imported by production.
 *
 * Font selection replaces the font. Symbol-wide bold is instead composed after
 * its deferred argument has been parsed, following MathJax's Boldsymbol token
 * postprocessor. The only new hook lives on command definitions, never on atoms,
 * MathML, or serialized LaTeX.
 */
function replaceExactly(
  source: string,
  target: string,
  replacement: string,
  label: string,
): string {
  const count = source.split(target).length - 1;
  if (count !== 1) {
    throw new Error("CP254 font-scope " + label + " anchor changed (" + count + ").");
  }
  return source.replace(target, replacement);
}

const symbolBoldDefinition = [
  'defineFunction(["boldsymbol", "bm", "bold"], "{:math*}", {',
  '  applyMode: "math",',
  '  applyStyle: (style) => __spreadProps(__spreadValues({}, style), { variantStyle: "bold" })',
  '});',
].join("\n");

const symbolBoldReplacement = [
  '// CP254: resolve symbol-wide weight after inner font selectors and resets.',
  '// atom.children is MathLive\'s complete descendant list, including branches.',
  '// A Set prevents repeated processing where roots/descendants overlap.',
  'function cp254ApplySymbolBoldToArgument(atoms) {',
  '  const targets = new Set();',
  '  for (const atom of atoms) {',
  '    targets.add(atom);',
  '    for (const child of atom.children) targets.add(child);',
  '  }',
  '  for (const atom of targets) {',
  '    if (atom.mode !== "math" || atom.type === "first") continue;',
  '    const style = atom.style || {};',
  '    const variant = style.variant;',
  '    const current = style.variantStyle;',
  '    let next;',
  '    if (variant === void 0 || variant === "normal" ||',
  '        variant === "main" || variant === "ams") {',
  '      // Preserve the existing implicit-family representation of symbol bold.',
  '      // Explicit italic resets acquire weight without losing their slant.',
  '      if (current === "bold" || current === "bolditalic") continue;',
  '      if (current === "italic") next = "bolditalic";',
  '      else if (current === void 0 || current === "up") next = "bold";',
  '      else continue;',
  '    } else if (variant === "calligraphic" || variant === "script" ||',
  '               variant === "fraktur" || variant === "sans-serif") {',
  '      // These are the nonbasic families in MathJax BOLDVARIANT.',
  '      // In particular, its table has no sans-serif-italic -> bold mapping.',
  '      if (current === void 0 || current === "up") next = "bold";',
  '      else continue;',
  '    } else {',
  '      // Double-struck, monospace, and unknown families have no supported',
  '      // weight projection in that authority; keep their exact final style.',
  '      continue;',
  '    }',
  '    atom.style = __spreadProps(__spreadValues({}, style), { variantStyle: next });',
  '    // Cached LaTeX/JSON was produced before this effective style existed.',
  '    // Use the existing invalidation path; add no AST provenance field.',
  '    atom.isDirty = true;',
  '  }',
  '  return atoms;',
  '}',
  'defineFunction(["boldsymbol", "bm", "bold"], "{:math*}", {',
  '  applyMode: "math",',
  '  applyStyle: (style) => style,',
  '  applyStyleToArgument: cp254ApplySymbolBoldToArgument',
  '});',
].join("\n");

export function patchCp254FontScope(source: string): string {
  if (!source.includes('getPropertyRuns(run, "visualTexBoldScope")')) {
    throw new Error("CP254 expects the current production-transformed owned kernel.");
  }
  if (source.includes("function cp254ApplySymbolBoldToArgument(")) {
    throw new Error("CP254 font-scope candidate was already applied.");
  }

  let patched = replaceExactly(
    source,
    [
      '    applyStyle: options.applyStyle,',
      '    serialize: options.serialize,',
    ].join("\n"),
    [
      '    applyStyle: options.applyStyle,',
      '    applyStyleToArgument: options.applyStyleToArgument,',
      '    serialize: options.serialize,',
    ].join("\n"),
    "command-definition argument-style hook",
  );

  patched = replaceExactly(
    patched,
    [
      '          const atoms = this.scanArgument(deferredArg);',
      '          this.style = saveStyle;',
      '          this.parseMode = savedMode2;',
      '          return argAtoms(atoms);',
    ].join("\n"),
    [
      '          const atoms = this.scanArgument(deferredArg);',
      '          this.style = saveStyle;',
      '          this.parseMode = savedMode2;',
      '          const resultAtoms = argAtoms(atoms);',
      '          return typeof info.applyStyleToArgument === "function"',
      '            ? info.applyStyleToArgument(resultAtoms)',
      '            : resultAtoms;',
    ].join("\n"),
    "deferred argument completion",
  );

  patched = replaceExactly(
    patched,
    symbolBoldDefinition,
    symbolBoldReplacement,
    "symbol-bold command family",
  );

  // These font selectors are replacement operations in MathJax MathFont, just
  // like mathsf, mathrm, and mathit. The previous removeItalic() accidentally
  // retained a basic outer mathbf weight. Symbol-wide bold now composes later.
  for (const family of ["double-struck", "fraktur", "calligraphic", "script"]) {
    patched = replaceExactly(
      patched,
      [
        '    variant: "' + family + '",',
        '    variantStyle: removeItalic(style.variantStyle)',
      ].join("\n"),
      [
        '    variant: "' + family + '",',
        '    variantStyle: "up"',
      ].join("\n"),
      family + " font replacement",
    );
  }
  patched = replaceExactly(
    patched,
    '    if (weightString(x[0].parent) === "bold")',
    [
      '    if (weightString(x[0].parent) === "bold" &&',
      '        (x[0].style.variant === "normal" && x[0].style.variantStyle === "bold") ===',
      '        (x[0].parent.style.variant === "normal" && x[0].parent.style.variantStyle === "bold"))',
    ].join("\n"),
    "keep font-scope reset inside symbol-bold parent",
  );
  return patched;
}

/**
 * Inputs for the owner's existing production editor/MathJax harness.
 * Expected variants describe the A token after authoritative MathJax export.
 * These are control data, not executed tests or acceptance claims.
 */
export const cp254FontScopeControls = [
  { source: String.raw`\mathbf{\mathsf{A}}`, expected: "sans-serif" },
  { source: String.raw`\boldsymbol{\mathsf{A}}`, expected: "bold-sans-serif" },
  { source: String.raw`\mathsf{\boldsymbol{A}}`, expected: "bold-sans-serif" },
  { source: String.raw`\mathbf{\mathfrak{A}}`, expected: "fraktur" },
  { source: String.raw`\boldsymbol{\mathfrak{A}}`, expected: "bold-fraktur" },
  { source: String.raw`\mathfrak{\boldsymbol{A}}`, expected: "bold-fraktur" },
  { source: String.raw`\boldsymbol{\mathsf{\mathfrak{A}}}`, expected: "bold-fraktur" },
  { source: String.raw`\boldsymbol{\mathfrak{\mathsf{A}}}`, expected: "bold-sans-serif" },
  { source: String.raw`\boldsymbol{\mathrm{A}}`, expected: "bold" },
  { source: String.raw`\boldsymbol{\mathnormal{A}}`, expected: "bold-italic" },
  { source: String.raw`\mathbf{\mathnormal{A}}`, expected: "italic" },
  { source: String.raw`\boldsymbol{\mathbb{A}}`, expected: "double-struck" },
  { source: String.raw`\boldsymbol{\mathtt{A}}`, expected: "monospace" },
  { source: String.raw`\boldsymbol{\boldsymbol{\mathsf{A}}}`, expected: "bold-sans-serif" },
] as const;
