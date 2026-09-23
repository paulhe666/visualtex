/** Pointer and selection geometry changes compiled into MathLive 0.109.2. */
export function patchVisualTexMathLiveSelection(source: string) {
  const replace = (target: string, replacement: string) => {
    const count = source.split(target).length - 1;
    if (count !== 1) {
      throw new Error(
        `MathLive selection anchor changed (${count}): ${target.slice(0, 80)}`,
      );
    }
    source = source.replace(target, replacement);
  };

  // Range aggregation must never enlarge the cached rectangle of its first atom.
  const rangeBoundsAnchor = `function getRangeBoundingRect(mf, range2) {
  const [start, end] = range2;
  let result = null;
  for (let i = start; i <= end; i++) {
    const bounds = getAtomBounds(mf, mf.model.at(i));
    if (bounds) {
      if (!result) result = bounds;`;
  replace(
    rangeBoundsAnchor,
    rangeBoundsAnchor.replace("result = bounds;", "result = { ...bounds };"),
  );

  replace(
    '`-${atom.parentBranch[0]}/${atom.parentBranch[0]}`',
    '`-${atom.parentBranch[0]}/${atom.parentBranch[1]}`',
  );

  // Accent atoms keep MathLive's kernel-level atomic navigation semantics:
  // skipBoundary=true and captureSelection=true. VisualTeX must not reopen
  // internal caret stops for a rendered accent after its argument is committed.
  // Keep the native deleteRange promotion for accents. Selecting/deleting the
  // complete accent argument promotes the range to the AccentAtom, so the
  // accent mark and argument are removed as one structural unit.

  replace(
    `  for (const atom of mathfield.model.getAtoms(range2, {
    includeChildren: true
  })) {`,
    `  const selectedAtoms = mathfield.model.getAtoms(range2, { includeChildren: true });
  const selectedSet = new Set(selectedAtoms);
  for (const atom of selectedAtoms) {
    // Sentinels have the surrounding font's ascent, not visible ink. Including
    // them creates tall phantom highlights in scripts and matrix cells.
    if (atom.type === "first") continue;`,
  );
  replace(
    "      const id = branchId(atom);",
    `      // Draw a fully selected structure once, including all protruding scripts.
      // Partial branches/cells keep separate rectangles and cannot tint siblings.
      let owner = atom;
      while (owner.parent && selectedSet.has(owner.parent)) owner = owner.parent;
      const id = branchId(owner);`,
  );

  replace(
    "      while ((_a3 = model.at(pos + dir)) == null ? void 0 : _a3.isDigit()) pos += dir;",
    `      while (true) {
        const next = model.at(pos + dir);
        if (!next?.isDigit() || next.parent !== atom.parent || next.parentBranch !== atom.parentBranch)
          break;
        pos += dir;
      }`,
  );

  const styleStart = source.indexOf(
    '    if (atom.style.variant || atom.style.variantStyle) {',
    source.indexOf("function boundary(model,"),
  );
  const styleEnd = source.indexOf("    return pos;", styleStart);
  if (styleStart < 0 || styleEnd < 0) {
    throw new Error("MathLive math group selection anchor changed");
  }
  // Math variables are atoms, not words. Typography does not define a group.
  source = source.slice(0, styleStart) + source.slice(styleEnd);

  replace(
    `          trackingWords = true;
          selectGroup(mathfield.model);`,
    `          trackingWords = true;
          // A caret offset on the glyph's left half belongs to its left sibling.
          // Double click selects the pointed atom, independently of that bias.
          const inkContainsPoint = node => {
            let r = node.getBoundingClientRect();
            const svg = node.querySelector("svg");
            if (svg) {
              const box = svg.getBBox(), matrix = svg.getScreenCTM();
              if (matrix && box.width && box.height) {
                const a = new DOMPoint(box.x, box.y).matrixTransform(matrix);
                const b = new DOMPoint(box.x + box.width, box.y + box.height).matrixTransform(matrix);
                r = {
                  left: a.x,
                  right: b.x,
                  top: a.y,
                  bottom: b.y,
                  width: b.x - a.x,
                  height: b.y - a.y,
                };
              }
            }
            return r.width > 0 && r.height > 0 &&
              pointInRect(evt.clientX, evt.clientY, r);
          };
          let ink = evt.composedPath().find(node =>
            node instanceof Element &&
            ["ML__visualtex-accent-mark", "ML__sqrt-sign", "ML__frac-line", "ML__open", "ML__close"]
              .some(name => node.classList.contains(name)) &&
            inkContainsPoint(node)
          );
          // Tall accents may paint above the content box: the browser targets
          // the container there. Resolve their visible mark before clamping Y.
          const directNode = evt.composedPath().find(node =>
            node instanceof Element && node.hasAttribute("data-atom-id")
          );
          const directAtom = directNode &&
            mathfield.model.root.children.find(
              atom => atom.id === directNode.getAttribute("data-atom-id"),
            );
          if (!ink && !directAtom?.value) {
            ink = [...mathfield.field.querySelectorAll(
              ".ML__visualtex-accent-mark, .ML__sqrt-sign, .ML__frac-line",
            )].find(inkContainsPoint);
          }
          const ownerId = ink?.closest("[data-atom-id]")?.getAttribute("data-atom-id");
          const owner = ownerId
            ? mathfield.model.root.children.find(atom => atom.id === ownerId)
            : null;
          const pointedOffset = owner
            ? mathfield.model.offsetOf(owner)
            : offsetFromPoint(
                mathfield,
                selectionAnchorX,
                selectionAnchorY,
                { bias: 1 },
              );
          const pointedAtom = mathfield.model.at(pointedOffset);
          if (pointedAtom?.hasChildren) {
            mathfield.model.setSelection(
              Math.max(0, mathfield.model.offsetOf(pointedAtom.leftSibling)),
              pointedOffset,
            );
          } else {
            if (pointedOffset >= 0) mathfield.model.position = pointedOffset;
            selectGroup(mathfield.model);
          }`,
  );

  // Placeholder hit testing must be authoritative before MathLive clamps a
  // pointer back into the field box. A denominator/script placeholder may
  // visibly extend outside that box even though its own atom bounds are valid.
  replace(
    `  if (gLastTap && Math.abs(gLastTap.x - anchorX) < 5 && Math.abs(gLastTap.y - anchorY) < 5 && Date.now() < gLastTap.time + 500) {`,
    `  const visualTexExactPlaceholderNode = [
    ...mathfield.field.querySelectorAll(
      ".visualtex-structural-placeholder[data-atom-id], .ML__placeholder[data-atom-id]"
    )
  ].find((node) => pointInRect(anchorX, anchorY, node.getBoundingClientRect()));
  const visualTexExactPlaceholderId =
    visualTexExactPlaceholderNode?.getAttribute("data-atom-id");
  const visualTexExactPlaceholder = visualTexExactPlaceholderId
    ? that.model.atoms.find((atom) => atom.id === visualTexExactPlaceholderId)
    : null;
  const visualTexExactPlaceholderOffset = visualTexExactPlaceholder
    ? that.model.offsetOf(visualTexExactPlaceholder)
    : -1;
  if (
    !evt.shiftKey &&
    visualTexExactPlaceholderOffset >= 0 &&
    (visualTexExactPlaceholder.type === "placeholder" ||
      visualTexExactPlaceholder.type === "prompt")
  ) {
    anchor = visualTexExactPlaceholderOffset;
    that.flushInlineShortcutBuffer();
    that.model.setSelection(
      Math.max(0, visualTexExactPlaceholderOffset - 1),
      visualTexExactPlaceholderOffset
    );
    mathfield.element.classList.add("tracking");
    trackingPointer = true;
    PointerTracker.start(field, evt, onPointerMove, endPointerTracking);
    if (!mathfield.hasFocus()) {
      mathfield.onFocus();
      mathfield.model.announce("line");
    }
    mathfield.stopCoalescingUndo();
    requestUpdate(mathfield);
    evt.preventDefault();
    return;
  }
  if (gLastTap && Math.abs(gLastTap.x - anchorX) < 5 && Math.abs(gLastTap.y - anchorY) < 5 && Date.now() < gLastTap.time + 500) {`,
  );

  replace(
    "function onPointerDown(mathfield, evt) {",
    `
function visualTexPointerRange(model, range) {
  let [start, end] = range;
  if (start === end) return range;
  let firstIndex = start + 1;
  while (firstIndex < end && model.at(firstIndex)?.type === "first") firstIndex++;
  let first = model.at(firstIndex), last = model.at(end);
  if (!first || !last) return range;
  const parent = Atom.commonAncestor(first, last);
  if (!parent) return range;
  while (first.parent && first.parent !== parent) first = first.parent;
  while (last.parent && last.parent !== parent) last = last.parent;
  // Crossing numerator/denominator or upper/lower scripts selects that whole
  // structure. Cells remain separate so a matrix row can be selected locally.
  const sameBranch =
    JSON.stringify(first.parentBranch) === JSON.stringify(last.parentBranch);
  if (!sameBranch && parent.type !== "array" && parent.type !== "root") {
    let before = parent.leftSibling;
    if (parent.type === "subsup" && before?.leftSibling) before = before.leftSibling;
    return [Math.max(0, model.offsetOf(before)), model.offsetOf(parent)];
  }
  // A selection entering/leaving a nested structure takes that structure,
  // never a flattened list of its children with the radical/fence discarded.
  start = Math.min(start, Math.max(0, model.offsetOf(first.leftSibling)));
  end = Math.max(end, model.offsetOf(last));
  return [start, end];
}
function onPointerDown(mathfield, evt) {`,
  );

  replace(
    `      const range2 = this.normalizeRange([anchor, position]);
      let [start, end] = range2;`,
    `      const range2 = this.normalizeRange([anchor, position]);
      let [start, end] = visualTexPointerRange(this, range2);`,
  );

  replace(
    `        "genfrac",
        "subsup",
        "accent",`,
    `        "genfrac",
        "surd",
        "leftright",
        "subsup",
        "accent",`,
  );

  return source;
}
