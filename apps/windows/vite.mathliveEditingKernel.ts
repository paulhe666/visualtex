/** Model commands compiled into the MathLive module, with no React/DOM dependency. */
export function patchVisualTexMathLiveEditingKernel(source: string) {
  const anchor = "// src/editor/undo.ts";
  if (source.split(anchor).length !== 2) {
    throw new Error("VisualTeX MathLive editing kernel anchor changed");
  }
  return source.replace(anchor, String.raw`
// VisualTeX: edit the existing atom tree instead of serializing and reparsing it.
function visualTexDefaultItalic(atom, mathfield) {
  const shape = atom.style.letterShapeStyle || mathfield.letterShapeStyle || "tex";
  return LETTER_SHAPE_RANGES.some((range, i) =>
    range.test(atom.value || "") && LETTER_SHAPE_MODIFIER[shape][i] === "it");
}
function visualTexAtomFont(atom, mathfield) {
  const style = atom.style;
  if (atom.mode === "text") return {
    bold: style.fontSeries === "b" || style.fontSeries === "bx",
    italic: style.fontShape === "it" || style.fontShape === "sl"
  };
  const variant = style.variant || "normal";
  return {
    bold: (style.variantStyle || "").includes("bold"),
    italic: (style.variantStyle || "").includes("italic") ||
      (!style.variantStyle && variant === "normal" && visualTexDefaultItalic(atom, mathfield))
  };
}
const VISUALTEX_PERSISTENT_STYLE_KEYS = [
  "variant",
  "variantStyle",
  "fontSeries",
  "fontShape",
  "color",
  "backgroundColor",
  "verbatimColor",
  "verbatimBackgroundColor"
];
function visualTexPersistentTypingGlyphs(atoms) {
  const result = [];
  const seen = new Set();
  const visit = (atom) => {
    if (!atom || typeof atom !== "object" || seen.has(atom)) return;
    seen.add(atom);
    const children = Array.isArray(atom.children) ? atom.children : [];
    if (
      children.length === 0 &&
      atom.mode !== "latex" &&
      atom.type !== "first" &&
      atom.type !== "placeholder" &&
      atom.type !== "prompt" &&
      typeof atom.value === "string" &&
      atom.value.length > 0
    ) {
      result.push(atom);
    }
    for (const child of children) visit(child);
  };
  for (const atom of atoms || []) visit(atom);
  return result;
}
function visualTexPersistentStyleSnapshot(atom) {
  const result = {};
  const style = atom.style || {};
  for (const key of VISUALTEX_PERSISTENT_STYLE_KEYS) {
    if (Object.prototype.hasOwnProperty.call(style, key))
      result[key] = style[key];
  }
  return result;
}
function visualTexPersistentStyleValue(snapshot, key) {
  return Object.prototype.hasOwnProperty.call(snapshot, key)
    ? snapshot[key]
    : undefined;
}
function visualTexPersistentStyleSource(mathfield) {
  const model = mathfield.model;
  const bias = mathfield.styleBias;
  if (!model || bias === "none") return null;
  const atom = model.at(model.position);
  if (!atom) return null;
  const before = ungroup(model, atom, bias);
  const after = ungroup(model, atom.rightSibling, bias);
  const offset = bias === "right" ? after : before;
  if (!Number.isFinite(offset) || offset < 0) return null;
  return model.at(offset);
}
function visualTexRestoreInheritedPersistentStyle(mathfield, atoms) {
  const source = visualTexPersistentStyleSource(mathfield);
  const marker = source && source.visualTexPersistentTypingApplied;
  if (!marker || !Array.isArray(marker.keys) || marker.keys.length === 0) return;
  for (const atom of visualTexPersistentTypingGlyphs(atoms)) {
    const style = { ...atom.style };
    let changed = false;
    for (const key of marker.keys) {
      const inherited = visualTexPersistentStyleValue(marker.after, key);
      const current = Object.prototype.hasOwnProperty.call(style, key)
        ? style[key]
        : undefined;
      if (current !== inherited) continue;
      const baseline = visualTexPersistentStyleValue(marker.before, key);
      if (baseline === undefined) delete style[key];
      else style[key] = baseline;
      changed = true;
    }
    if (changed) {
      atom.style = style;
      atom.isDirty = true;
    }
  }
}
function visualTexApplyPersistentFontStyle(atom, mathfield, persistent) {
  if (!persistent.bold && !persistent.italic) return;
  const font = visualTexAtomFont(atom, mathfield);
  if (persistent.bold) font.bold = true;
  if (persistent.italic) font.italic = true;
  const style = { ...atom.style };
  if (atom.mode === "text") {
    if (persistent.bold) style.fontSeries = font.bold ? "b" : "m";
    if (persistent.italic) style.fontShape = font.italic ? "it" : "n";
  } else if (atom.mode === "math") {
    const variant = style.variant || "normal";
    if (!["normal", "main", "ams"].includes(variant)) {
      style.variantStyle = font.bold ? (font.italic ? "bolditalic" : "bold") :
        (font.italic ? "italic" : undefined);
    } else {
      style.variant = font.italic ? "main" : "normal";
      style.variantStyle = font.bold ? (font.italic ? "bolditalic" : "bold") :
        (font.italic ? "italic" : "up");
      if (!font.bold && font.italic === visualTexDefaultItalic(atom, mathfield)) {
        style.variant = "normal";
        style.variantStyle = undefined;
      }
    }
  }
  atom.style = style;
  atom.isDirty = true;
}
function visualTexApplyPersistentTypingStyle(mathfield, atoms) {
  // Native MathLive inherits style from the glyph next to the caret. First
  // remove only the style delta that VisualTeX itself added to that source
  // glyph; this prevents a disabled persistent style from leaking forward
  // without breaking normal \mathbf/\mathrm/color inheritance.
  visualTexRestoreInheritedPersistentStyle(mathfield, atoms);

  const persistent = mathfield.visualTexPersistentTypingStyle;
  if (!persistent || (
    !persistent.bold &&
    !persistent.italic &&
    !persistent.color &&
    !persistent.backgroundColor
  )) return;

  for (const atom of visualTexPersistentTypingGlyphs(atoms)) {
    const before = visualTexPersistentStyleSnapshot(atom);
    visualTexApplyPersistentFontStyle(atom, mathfield, persistent);
    const style = { ...atom.style };
    if (persistent.color) style.color = persistent.color;
    if (persistent.backgroundColor) style.backgroundColor = persistent.backgroundColor;
    atom.style = style;
    atom.isDirty = true;

    const after = visualTexPersistentStyleSnapshot(atom);
    const changedKeys = VISUALTEX_PERSISTENT_STYLE_KEYS.filter(
      (key) =>
        visualTexPersistentStyleValue(before, key) !==
        visualTexPersistentStyleValue(after, key),
    );
    if (changedKeys.length > 0) {
      atom.visualTexPersistentTypingApplied = {
        before,
        after,
        keys: changedKeys
      };
    } else {
      delete atom.visualTexPersistentTypingApplied;
    }
  }
}
function visualTexToggleSelectionStyle(model, kind, requestedSelection) {
  const selection = requestedSelection ? model.normalizeSelection(requestedSelection) : model.selection;
  if (kind !== "bold" && kind !== "italic" || !selection ||
      selection.ranges.every(([start, end]) => start === end)) return false;
  // Public setSelection() collapses disjoint ranges to their enclosing range.
  // Accept the caller's explicit ranges without formatting the gaps between them.
  const atoms = [...new Set(model.getAtoms(selection, { includeChildren: true }))];
  // Containers (accent, fraction, root...) are structural, not font-bearing glyphs.
  const glyphs = atoms.filter(atom => atom.type !== "first" && atom.value &&
    atom.type !== "placeholder" && atom.type !== "prompt");
  if (!glyphs.length) return false;
  const inputType = kind === "bold" ? "formatBold" : "formatItalic";
  if (!model.contentWillChange({ inputType })) return false;
  const mathfield = model.mathfield;
  mathfield.flushInlineShortcutBuffer();
  mathfield.stopCoalescingUndo();
  if (mathfield.undoManager.index < 0) mathfield.snapshot();
  const fonts = glyphs.map(atom => visualTexAtomFont(atom, mathfield));
  const enabled = !fonts.every(font => font[kind]);
  model.deferNotifications({ content: true, type: inputType }, () => {
    // Parsed font commands copy style onto their descendants. Clear selected
    // container styles too, otherwise their old context overrides a glyph reset.
    for (const atom of atoms) {
      if (glyphs.includes(atom) || atom.type === "first") continue;
      atom.style = { ...atom.style, variant: undefined, variantStyle: undefined };
      atom.isDirty = true;
    }
    glyphs.forEach((atom, i) => {
      const font = { ...fonts[i], [kind]: enabled };
      const style = { ...atom.style };
      if (atom.mode === "text") {
        style.fontSeries = font.bold ? "b" : "m";
        style.fontShape = font.italic ? "it" : "n";
      } else {
        const variant = style.variant || "normal";
        if (!["normal", "main", "ams"].includes(variant)) {
          // Preserve a selected alphabet (mathbb, mathfrak, etc.).
          style.variantStyle = font.bold ? (font.italic ? "bolditalic" : "bold") :
            (font.italic ? "italic" : undefined);
        } else {
          style.variant = font.italic ? "main" : "normal";
          style.variantStyle = font.bold ? (font.italic ? "bolditalic" : "bold") :
            (font.italic ? "italic" : "up");
          if (!font.bold && font.italic === visualTexDefaultItalic(atom, mathfield)) {
            style.variant = "normal";
            style.variantStyle = undefined;
          }
        }
      }
      atom.style = style;
      atom.isDirty = true;
    });
    mathfield.snapshot();
  });
  return true;
}
function visualTexAlignedCellIsEmpty(cell) {
  return !cell || cell.every(atom => atom.type === "first");
}
function visualTexProvisionalAlignRows(array) {
  if (!(array.visualTexProvisionalAlignRows instanceof Set)) {
    array.visualTexProvisionalAlignRows = new Set();
  }
  return array.visualTexProvisionalAlignRows;
}
function visualTexEnsureAlignedColumnCapacity(array, requiredColumns) {
  while (array.maxColumns < requiredColumns) {
    const trailing = array.colFormat[array.colFormat.length - 1];
    const hasTrailingGap = trailing && Object.prototype.hasOwnProperty.call(trailing, "gap");
    if (hasTrailingGap) array.colFormat.pop();
    array.colFormat.push(
      { gap: 1 },
      { align: "r" },
      { gap: 0.25 },
      { align: "l" },
      hasTrailingGap ? trailing : { gap: 0 },
    );
    array.isDirty = true;
  }
}
function visualTexInsertAlignmentPoint(model) {
  const array = model.parentEnvironment;
  if (!array || array.type !== "array" ||
      !["aligned", "align", "align*"].includes(array.environmentName)) return false;
  if (!model.contentWillChange({ inputType: "insertText" })) return false;

  const mathfield = model.mathfield;
  mathfield.flushInlineShortcutBuffer();
  mathfield.stopCoalescingUndo();
  if (mathfield.undoManager.index < 0) mathfield.snapshot();

  model.deferNotifications({ content: true, selection: true, type: "insertText" }, () => {
    if (!model.selectionIsCollapsed) model.deleteAtoms(range(model.selection));
    const cursor = model.at(model.position);
    if (cursor.parent !== array || !Array.isArray(cursor.parentBranch)) return;
    const [row, col] = cursor.parentBranch;
    const siblings = array.getCell(row, col);
    if (!siblings) return;
    const split = siblings.indexOf(cursor) + 1;
    const before = siblings.slice(1, split);
    const after = siblings.slice(split);
    const provisionalRows = visualTexProvisionalAlignRows(array);

    // A freshly-created VisualTeX align row starts in MathLive's left-aligned
    // second cell so typing grows naturally from left to right. The first '&'
    // commits that provisional row to standard TeX align semantics by moving
    // the content before the caret into the right-aligned first cell.
    if (
      col === 1 &&
      visualTexAlignedCellIsEmpty(array.getCell(row, 0)) &&
      (
        provisionalRows.has(row) ||
        (array.environmentName === "aligned" && array.colCount === 2)
      )
    ) {
      array.setCell(row, 0, before);
      array.setCell(row, 1, after);
      provisionalRows.delete(row);
      model.position = model.offsetOf(array.getCell(row, 1)[0]);
      mathfield.snapshot();
      return;
    }

    const targetCol = col + 1;
    const targetCell = array.getCell(row, targetCol);

    array.setCell(row, col, before);
    if (!targetCell || !visualTexAlignedCellIsEmpty(targetCell)) {
      visualTexEnsureAlignedColumnCapacity(array, array.colCount + 1);
      array.addColumnAfter(col);
    }
    array.setCell(row, targetCol, after);
    model.position = model.offsetOf(array.getCell(row, targetCol)[0]);
    mathfield.snapshot();
  });
  return true;
}
function visualTexInsertRowBreak(model, environment = "gathered") {
  if (environment !== "aligned" && environment !== "gathered") return false;
  if (!model.contentWillChange({ inputType: "insertLineBreak" })) return false;
  const mathfield = model.mathfield;
  mathfield.flushInlineShortcutBuffer();
  mathfield.stopCoalescingUndo();
  if (mathfield.undoManager.index < 0) mathfield.snapshot();
  model.deferNotifications({ content: true, selection: true, type: "insertLineBreak" }, () => {
    if (!model.selectionIsCollapsed) model.deleteAtoms(range(model.selection));
    let cursor = model.at(model.position);
    const body = model.root.body?.filter(atom => atom.type !== "first");
    // A click outside a standalone array is a root position, not a cell position.
    if (body?.length === 1 && body[0].type === "array" &&
        (model.position === 0 || model.position === model.lastOffset)) {
      const array = body[0];
      const atStart = model.position === 0;
      setPositionInCell(model, array, atStart ? 0 : array.rowCount - 1,
        atStart ? 0 : array.colCount - 1, atStart ? "start" : "end");
      cursor = model.at(model.position);
    }
    // Split the current branch. At a cell boundary this extends the array;
    // inside a fraction/accent it creates local lines without flattening it.
    const parent = cursor.parent;
    const branch = cursor.parentBranch;
    if (!parent || branch == null) return;
    if (Array.isArray(branch)) {
      const [row, col] = branch;
      const siblings = parent.getCell(row, col);
      const split = siblings.indexOf(cursor) + 1;
      const before = siblings.slice(1, split);
      const after = siblings.slice(split);
      parent.setCell(row, col, before);
      parent.addRowAfter(row);
      if (parent.environmentName === "aligned") {
        for (let column = 0; column < parent.colCount; column += 1) {
          parent.setCell(row + 1, column, []);
        }
        // Put a newly-created align row into the left-aligned second cell.
        // This keeps normal typing direction intuitive until the user enters
        // an explicit '&', which visualTexInsertAlignmentPoint then commits
        // into the standard right/left TeX column pair.
        visualTexProvisionalAlignRows(parent).add(row + 1);
        parent.setCell(row + 1, 1, after);
        model.position = model.offsetOf(parent.getCell(row + 1, 1)[0]);
      } else {
        parent.setCell(row + 1, col, after);
        model.position = model.offsetOf(parent.getCell(row + 1, col)[0]);
      }
    } else {
      const siblings = parent.branch(branch);
      const split = siblings.indexOf(cursor) + 1;
      const before = siblings.slice(1, split);
      const after = siblings.slice(split);
      parent.removeBranch(branch);
      const array =
        environment === "aligned"
          ? makeEnvironment(environment, [[[], before], [[], after]])
          : makeEnvironment(environment, [[before], [after]]);
      if (environment === "aligned") {
        const provisionalRows = visualTexProvisionalAlignRows(array);
        provisionalRows.add(0);
        provisionalRows.add(1);
      }
      parent.setChildren([array], branch);
      model.position = model.offsetOf(
        array.getCell(1, environment === "aligned" ? 1 : 0)[0],
      );
    }
    mathfield.snapshot();
  });
  return true;
}
register2({
  visualTexToggleSelectionStyle,
  visualTexInsertAlignmentPoint,
  visualTexInsertRowBreak,
}, {
  target: "model", canUndo: true, changeContent: true, changeSelection: true
});

` + anchor);
}
