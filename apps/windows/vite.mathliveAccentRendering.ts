/** Accent geometry is generated with MathLive's Box tree, before DOM creation. */
export function patchVisualTexMathLiveAccentRendering(source: string) {
  const replace = (target: string, replacement: string) => {
    const count = source.split(target).length - 1;
    if (count !== 1) throw new Error(`MathLive accent render anchor changed (${count})`);
    source = source.replace(target, replacement);
  };
  replace("// src/atoms/accent.ts", String.raw`
// Native vector/dot ink: combining Unicode glyphs have zero advance and depend
// on the surrounding font. Give their ink a real box, also for static rendering.
function visualTexCombiningAccentBox(codepoint) {
  const vector = codepoint === 8407;
  const count = codepoint === 8411 ? 3 : 4;
  const box = new Box(null, { classes: "ML__accent-body ML__visualtex-accent-ink" });
  box.width = vector ? 0.42 : count === 3 ? 0.41 : 0.57;
  // Retain the ascent/descent of the combining glyph's native TeX box.
  box.height = vector ? 0.72 : 0.70;
  box.depth = vector ? 0 : 0.20;
  box.setStyle("position", "relative");
  const y = (box.height + box.depth) / 2;
  box.svgOverlay = vector
    ? '<path d="M0.015 0.36H0.39M0.31 0.28L0.39 0.36L0.31 0.44" fill="none" stroke="currentColor" stroke-width="0.035" stroke-linecap="round" stroke-linejoin="round"/>'
    : Array.from({ length: count }, (_, i) =>
        '<circle cx="' + (0.045 + i * 0.16) + '" cy="' + y + '" r="0.045" fill="currentColor"/>').join("");
  return box;
}
function visualTexMarkAccentLayout(box) {
  if (box.classes.split(" ").includes("ML__vlist")) {
    box.classes += " visualtex-accent-placeholder-layout";
    return true;
  }
  return (box.children || []).some(visualTexMarkAccentLayout);
}
// src/atoms/accent.ts`);
  replace(
    '    } else if (this.accent) {\n      const accent = new Box(this.accent, { fontFamily: "Main-Regular" });',
    '    } else if ([8407, 8411, 8412].includes(this.accent)) {\n      accentBox = visualTexCombiningAccentBox(this.accent);\n    } else if (this.accent) {\n      const accent = new Box(this.accent, { fontFamily: "Main-Regular" });',
  );
  replace(
    '    const accentMargin = (base.width - accentBox.width) / 2 + base.left + 2 * skew;',
    [
      '    // The center row already centers against the actual rendered base.',
      '    // A second metric-based half-width offset drifts with placeholder/font widths.',
      '    const accentMargin = 2 * skew;',
      '    if (this.svgAccent) accentBox.setStyle("width", "100%");',
    ].join("\n"),
  );
  replace(
    '    const result = new Box(accentBox, { type: "ord" });',
    [
      '    if (this.body?.length === 2 && this.body[1].type === "placeholder")',
      '      visualTexMarkAccentLayout(accentBox);',
      '    const result = new Box(accentBox, { type: "ord", classes: "ML__visualtex-accent" });',
    ].join("\n"),
  );
  replace(
    '    if (this.caret) result.classes += " ML__placeholder-selected";',
    [
      '    if (this.parent?.type === "accent" && this.parentBranch === "body" && this.parent.body.length === 2)',
      '      result.classes += " visualtex-accent-structural-placeholder";',
      '    if (this.caret) result.classes += " ML__placeholder-selected";',
    ].join("\n"),
  );
  return source;
}
