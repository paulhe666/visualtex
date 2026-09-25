import { commandRegistry } from "./src/autocomplete/commandRegistry";

/** Compile discovery, preview and completion into MathLive's own command UI. */
export function patchVisualTexMathLiveSemanticCompletion(source: string) {
  const replace = (from: string, to: string) => {
    if (source.split(from).length !== 2) throw new Error("MathLive semantic completion anchor changed: " + from.slice(0, 70));
    source = source.replace(from, to);
  };
  const records = commandRegistry.map(c => ({
    id: c.id, command: c.command, latex: c.insertTemplate, preview: c.previewLatex,
    // Search can still use translated keywords, but the native list only
    // presents the command and its mathematical preview.
    label: "", words: [c.labelEn, ...c.aliases, ...c.keywords], priority: c.defaultPriority,
  }));
  const start = source.indexOf("function suggest(mf, s) {");
  const end = source.indexOf("function parseParameterTemplateArgument", start);
  if (start < 0 || end < 0) throw new Error("MathLive suggest anchor changed");
  source = source.slice(0, start) + `const visualTexCompletionCatalog = ${JSON.stringify(records)};\n` + String.raw`
function visualTexCompletionInfo(mf, latex) {
  return mf.visualTexCompletionRecords?.get(latex) || {
    id: "mathlive-native:" + latex,
    command: latex,
    preview: latex,
    label: "",
  };
}
function visualTexPersonalScore(preferences, usageId, query) {
  if (!preferences?.personalize || !usageId) return 0;
  const usage = preferences.usage?.[usageId];
  if (!usage) return 0;

  const now = Date.now();
  const useCount = Math.max(0, usage.useCount || 0);
  const prefixUses = Math.max(0, usage.acceptedPrefixes?.[query] || 0);
  const recentUses = Array.isArray(usage.recentUses) ? usage.recentUses : [];

  // Lifetime frequency is deliberately logarithmic: long-term habits matter,
  // but they must not permanently pin an old command above a command the user
  // has started using heavily today.
  const lifetimeScore = Math.min(180, Math.log2(1 + useCount) * 34);
  const prefixScore = Math.min(140, Math.log2(1 + prefixUses) * 34);

  // Recent frequency is the strongest personalization signal. Each recent use
  // contributes independently with a smooth time decay rather than a hard
  // "N uses => force first" threshold.
  const recentScore = Math.min(
    420,
    recentUses.reduce((score, timestamp) => {
      const age = Math.max(0, now - timestamp);
      if (age <= 120_000) return score + 95;
      if (age <= 600_000) return score + 65;
      if (age <= 3_600_000) return score + 38;
      if (age <= 86_400_000) return score + 18;
      if (age <= 604_800_000) return score + 6;
      return score;
    }, 0),
  );

  const lastUsedAge = Math.max(0, now - (usage.lastUsedAt || 0));
  const lastUsedScore =
    lastUsedAge <= 120_000 ? 45 :
    lastUsedAge <= 600_000 ? 30 :
    lastUsedAge <= 3_600_000 ? 16 :
    lastUsedAge <= 86_400_000 ? 6 : 0;

  return (
    lifetimeScore +
    prefixScore +
    recentScore +
    lastUsedScore +
    (usage.pinned ? 160 : 0)
  );
}
function suggest(mf, s) {
  visualTexSyncSourceWrapper(mf);
  if (!s.startsWith("\\") || s.length < 2) return [];
  const query = s.slice(1).toLowerCase();
  const preferences = mf.host?.visualTexCompletionPreferences;
  const literalDelimiterCommands = new Set([
    "\\langle", "\\rangle",
    "\\lfloor", "\\rfloor",
    "\\lceil", "\\rceil",
    "\\lvert", "\\rvert",
    "\\lVert", "\\rVert",
    "\\vert", "\\Vert",
  ]);
  const exactLiteralDelimiter = literalDelimiterCommands.has(s);
  // \partial is an atomic MathLive symbol even though VisualTeX also exposes
  // a richer partial-derivative fraction template under the same discovery
  // name. An exact raw command must preserve the literal symbol; prefix/search
  // queries still keep the semantic derivative template available.
  const exactAtomicNativeCommands = new Set(["\\partial"]);
  const exactAtomicNative = exactAtomicNativeCommands.has(s);
  const results = new Map();
  const family = /^(int|integ)/.test(query) ? ["int-bare", "int", "oint", "intplain", "iint", "iiint"] :
    /^(sum|summ)/.test(query) ? ["sum", "series", "sigma", "Sigma"] :
    /^(lim|limit)/.test(query) ? ["lim", "lim-infty", "lim-left", "lim-right"] :
    /^(frac|fraction|divide)/.test(query) ? ["frac", "dfrac", "tfrac"] : [];
  const parameterizedCatalogCommands = new Set(
    visualTexCompletionCatalog
      .filter((record) => record.latex.includes("\\placeholder{}"))
      .map((record) => record.command),
  );
  const bareOperatorCommands = new Set([
    "\\int", "\\iint", "\\iiint", "\\oint", "\\oiint", "\\oiiint",
    "\\sum", "\\prod", "\\lim",
  ]);
  const add = (record, score) => {
    const old = results.get(record.latex);
    if (!old || score > old.score) results.set(record.latex, { ...record, score });
  };
  if (exactLiteralDelimiter) {
    add({
      id: "mathlive-literal:" + s,
      command: s,
      latex: s,
      preview: s,
      label: "",
    }, 6000);
  }
  for (const record of visualTexCompletionCatalog) {
    const name = record.command.replace(/^\\/, "");
    const normalizedName = name.toLowerCase();
    const rawQuery = s.slice(1);
    // When the user has typed a literal delimiter or an atomic native
    // symbol exactly, semantic templates must not replace that literal source.
    // Prefix discovery remains unchanged for non-exact queries.
    if (
      exactLiteralDelimiter ||
      (exactAtomicNative && record.command === s && record.latex !== s)
    ) continue;
    let score = -Infinity;
    if (name === rawQuery) score = 5000;
    else if (normalizedName === query) score = 760;
    else if (normalizedName.startsWith(query))
      score = 470 - Math.min(100, normalizedName.length - query.length);
    for (const word of record.words) {
      const normalized = word.toLowerCase();
      if (normalized === query) score = Math.max(score, 600);
      else if (normalized.startsWith(query)) score = Math.max(score, 450 - Math.min(100, normalized.length - query.length));
      else if (query.length >= 3 && normalized.split(/[^a-z0-9]+/).some(term => term.startsWith(query))) score = Math.max(score, 260);
    }
    const intent = family.indexOf(record.id);
    if (intent >= 0) score = 900 - intent * 25;
    if (!Number.isFinite(score)) continue;
    if (name !== rawQuery)
      score += visualTexPersonalScore(preferences, record.id, query);
    add(record, score + record.priority / 100);
  }
  // Native symbols and user macros remain available, with case-insensitive
  // discovery (del -> delta AND Delta) and case-sensitive exact preference.
  for (const command of [...Object.keys(LATEX_COMMANDS), ...Object.keys(MATH_SYMBOLS), ...Object.keys(mf.options.macros).map(x => "\\" + x)]) {
    if (!command.toLowerCase().startsWith(s.toLowerCase()) || LATEX_COMMANDS[command]?.infix) continue;
    // Parameterized commands such as \\ket, \\sqrt, \\vec, \\abs and
    // \\norm must never compete with their editable VisualTeX template as a
    // bare command. Large operators are the intentional exception: \\int,
    // \\sum, ... are meaningful by themselves and therefore keep both forms.
    if (
      parameterizedCatalogCommands.has(command) &&
      !bareOperatorCommands.has(command) &&
      !(exactAtomicNative && command === s)
    ) continue;
    const exact = command === s;
    const usageId = "mathlive-native:" + command;
    const score = (exact ? 5000 : 445 - (command.length - s.length)) +
      (exact ? 0 : visualTexPersonalScore(preferences, usageId, query));
    add({ id: usageId, command, latex: command, preview: command, label: "" }, score);
  }
  const ranked = [...results.values()].sort((a,b) => b.score-a.score || a.command.localeCompare(b.command));
  mf.visualTexCompletionRecords = new Map(ranked.map(r => [r.latex, r]));
  return ranked.slice(0, Math.max(6, Math.min(24, preferences?.count || 12))).map(r => r.latex);
}
` + source.slice(end);
  const ghostStart = source.indexOf("  const suggestion = suggestions[mathfield.suggestionIndex];", source.indexOf("function updateAutocomplete("));
  const ghostEnd = source.indexOf("  showSuggestionPopover(mathfield, suggestions);", ghostStart);
  if (ghostStart < 0 || ghostEnd < 0) throw new Error("MathLive inline suggestion anchor changed");
  // Semantic candidates may not share the typed prefix. Keep raw input intact;
  // their complete previews belong in the native panel, not in the formula.
  source = source.slice(0, ghostStart) + `  mathfield.visualTexCompletion = { command, suggestions, group: getLatexGroup(model) };\n` + source.slice(ghostEnd);
  replace(`  if (suggestions.length === 0) {
    if (/^\\\\[a-zA-Z\\*]+$/.test(command))`, `  if (suggestions.length === 0) {
    mathfield.visualTexCompletion = null;
    if (/^\\\\[a-zA-Z\\*]+$/.test(command))`);
  replace(`  hideSuggestionPopover(mathfield);
  const latexGroup = getLatexGroup(mathfield.model);`, `  const visualTexOwnedSourceGroup = mathfield.visualTexSourceWrapper?.group;
  const visualTexCurrentLatexGroup = getLatexGroup(mathfield.model);
  if (visualTexOwnedSourceGroup && visualTexCurrentLatexGroup === visualTexOwnedSourceGroup) {
    if (completion === "reject") return visualTexCancelSourceWrapper(mathfield);
    return false;
  }
  const completionState = mathfield.visualTexCompletion;
  const selectedCompletion = completionState?.suggestions[mathfield.suggestionIndex];
  hideSuggestionPopover(mathfield);
  const latexGroup = getLatexGroup(mathfield.model);`);
  replace(`  if (completion === "accept-suggestion" || completion === "accept-all") {`, `  const useSemanticCompletion = (completion === "accept-suggestion" || completion === "accept-all") && completionState?.group === latexGroup && selectedCompletion;
  if (!useSemanticCompletion && (completion === "accept-suggestion" || completion === "accept-all")) {`);
  replace(`  const latex = body.map((x) => x.value).join("");
  const newPos = latexGroup.leftSibling;`, `  const latex = useSemanticCompletion ? selectedCompletion : body.map((x) => x.value).join("");
  mathfield.visualTexCompletion = null;
  const newPos = latexGroup.leftSibling;`);
  replace(
    `  if (completion === "reject") return true;`,
    `  if (completion === "reject") {
    mathfield.visualTexPendingStructuralCommandSlot = null;
    return true;
  }`,
  );
  replace(`  ModeEditor.insert(mathfield.model, latex, {`, `  const visualTexAcceptedInfo = useSemanticCompletion
    ? visualTexCompletionInfo(mathfield, selectedCompletion)
    : null;
  const visualTexSourceWrapperCommand =
    visualTexAcceptedInfo?.command &&
    visualTexSourceWrapperKind(visualTexAcceptedInfo.command)
      ? visualTexCanonicalVariantInputCommand(visualTexAcceptedInfo.command)
      : null;
  if (visualTexSourceWrapperCommand) {
    const visualTexSourceSlot =
      mathfield.visualTexPendingStructuralCommandSlot || null;
    const visualTexStartedSourceWrapper = visualTexStartSourceWrapper(
      mathfield.model,
      visualTexSourceWrapperCommand + "{\\\\placeholder{}}",
      visualTexSourceSlot,
    );
    mathfield.snapshot("insert-source-wrapper");
    if (visualTexStartedSourceWrapper) {
      const info = visualTexAcceptedInfo;
      // MathLive contentDidChange() dispatches its public input event on a
      // zero-delay timer. Queue usage/acceptance after that timer so a
      // controlled host cannot rerender with the stale pre-completion value
      // before the model's input transaction has been observed.
      setTimeout(() => {
        if (!mathfield.host) return;
        mathfield.host.dispatchEvent(new CustomEvent("visualtex-command-accepted", {
          bubbles: true,
          composed: true,
          detail: {
            id: info.id || info.command.slice(1),
            query: completionState.command,
            command: info.command,
            latex: selectedCompletion,
            variantInput: true,
          },
        }));
      }, 0);
      mathfield.model.announce("replacement");
      return true;
    }
  }
  ModeEditor.insert(mathfield.model, latex, {`);
  replace(`  mathfield.snapshot();
  mathfield.model.announce("replacement");`, `  if (visualTexAcceptedInfo)
    visualTexSelectFirstSemanticWrapperPlaceholder(
      mathfield,
      visualTexAcceptedInfo.command,
    );
  mathfield.snapshot();
  if (useSemanticCompletion) {
    const info = visualTexAcceptedInfo;
    setTimeout(() => {
      if (!mathfield.host) return;
      mathfield.host.dispatchEvent(new CustomEvent("visualtex-command-accepted", {
        bubbles: true,
        composed: true,
        detail: {
          id: info.id || info.command.slice(1),
          query: completionState.command,
          command: info.command,
          latex: selectedCompletion,
          variantInput: false,
        },
      }));
    }, 0);
  }
  mathfield.model.announce("replacement");`);
  replace(`    const command = suggestion;
    const commandMarkup = latexToMarkup(mf, suggestion);`, `    const info = visualTexCompletionInfo(mf, suggestion);
    const command = suggestion;
    const commandMarkup = latexToMarkup(mf, info.preview);`);
  replace('${escapeHtmlAttr(command)}</span><span class="ML__popover__command">', '${escapeHtmlAttr(info.command)}${info.label ? " · " + escapeHtmlAttr(info.label) : ""}</span><span class="ML__popover__command">');
  replace(`const keybinding = getKeybindingsForCommand(mf.keybindings, command).join(`, `const keybinding = getKeybindingsForCommand(mf.keybindings, info.command).join(`);
  // Commands are literal LaTeX, including brackets and repeated backslashes.
  // Compare their prefix and command boundary without constructing a regexp.
  const keyStart = source.indexOf("  const regex = new RegExp(", source.indexOf("function getKeybindingsForCommand("));
  const keyEnd = source.indexOf("  for (const keybinding of keybindings)", keyStart);
  if (keyStart < 0 || keyEnd < 0) throw new Error("MathLive keybinding comparison anchor changed");
  source = source.slice(0, keyStart) + source.slice(keyEnd);
  replace(`    if (regex.test(commandToString(keybinding.command)))`, `    const candidateCommand = commandToString(keybinding.command);
    if (candidateCommand.startsWith(normalizedCommand) && !/[*a-zA-Z]/.test(candidateCommand.charAt(normalizedCommand.length)))`);
  // Keep one native panel alive: arrow navigation updates rows without a
  // remove/recreate cycle or an application-owned cloned popover.
  replace(`  if (panel) releaseSharedElement("mathlive-suggestion-popover");
}
function createSuggestionPopover`, `  if (panel?.visualTexMathfield === mf) { panel.classList.remove("is-visible"); panel.setAttribute("aria-hidden", "true"); }
}
function createSuggestionPopover`);
  const createStart = source.indexOf("function createSuggestionPopover(mf, html) {");
  const createEnd = source.indexOf("function disposeSuggestionPopover()", createStart);
  source = source.slice(0, createStart) + String.raw`
function createSuggestionPopover(mf, html) {
  let panel = document.getElementById("mathlive-suggestion-popover");
  if (!panel) {
    injectStylesheet("suggestion-popover"); injectStylesheet("core");
    panel = getSharedElement("mathlive-suggestion-popover");
    panel.addEventListener("pointerdown", ev => ev.preventDefault());
    panel.addEventListener("click", ev => {
      const item = ev.target.closest("li[data-command]");
      const owner = panel.visualTexMathfield;
      if (!item || !owner) return;
      const entries = owner.visualTexCompletion?.suggestions || [];
      const index = entries.indexOf(item.dataset.command);
      if (index < 0) return;
      owner.suggestionIndex = index;
      complete(owner, "accept-all"); owner.dirty = true; owner.focus();
      requestUpdate(owner);
    });
  }
  panel.visualTexMathfield = mf;
  panel.setAttribute("aria-hidden", "false");
  panel.innerHTML = globalThis.MathfieldElement.createHTML(html);
  return panel;
}
` + source.slice(createEnd);
  // A field may disconnect after another field has already opened the shared
  // panel. Only its current owner may dispose or reposition it.
  replace(`    disposeSuggestionPopover();`, `    disposeSuggestionPopover(this);`);
  replace(`function disposeSuggestionPopover() {
  if (!document.getElementById("mathlive-suggestion-popover")) return;`, `function disposeSuggestionPopover(mf) {
  const panel = document.getElementById("mathlive-suggestion-popover");
  if (!panel || panel.visualTexMathfield !== mf) return;`);
  replace(`    if (panel && !isSuggestionPopoverVisible()) {`, `    if (panel?.isConnected && panel.visualTexMathfield === mf && panel.getAttribute("aria-hidden") !== "true" && !isSuggestionPopoverVisible()) {`);
  replace(`  if (!isSuggestionPopoverVisible()) return;
  if (((_a3 = mf.model.at(mf.model.position))`, `  if (!isSuggestionPopoverVisible() || document.getElementById("mathlive-suggestion-popover")?.visualTexMathfield !== mf) return;
  if (((_a3 = mf.model.at(mf.model.position))`);
  // Keep command candidates inside the actual visible viewport. MathLive's
  // stock placement chooses above/below but can still produce a negative top
  // when the panel is taller than the available side, which clips the first
  // candidates near the top of the editor.
  replace(`  const spaceAbove = position.y - position.height;
  const spaceBelow = viewportHeight - scrollbarHeight - virtualkeyboardHeight - position.y;
  if (spaceBelow < spaceAbove) {
    panel.classList.add("ML__popover--reverse-direction");
    panel.classList.remove("top-tip");
    panel.classList.add("bottom-tip");
    panel.style.top = \`\${position.y - position.height - panel.offsetHeight - 15}px\`;
  } else {
    panel.classList.remove("ML__popover--reverse-direction");
    panel.classList.add("top-tip");
    panel.classList.remove("bottom-tip");
    panel.style.top = \`\${position.y + 15}px\`;
  }`, `  const viewportPadding = 8;
  const tipGap = 15;
  const viewportRight = viewportWidth - scrollbarWidth - viewportPadding;
  const viewportBottom =
    viewportHeight - scrollbarHeight - virtualkeyboardHeight - viewportPadding;
  const centeredLeft = position.x - panel.offsetWidth / 2;
  panel.style.left = \`\${Math.max(
    viewportPadding,
    Math.min(centeredLeft, viewportRight - panel.offsetWidth),
  )}px\`;

  const spaceAbove = Math.max(
    0,
    position.y - position.height - viewportPadding - tipGap,
  );
  const spaceBelow = Math.max(0, viewportBottom - position.y - tipGap);
  const placeAbove = spaceAbove > spaceBelow;
  const availableHeight = Math.max(
    36,
    Math.min(320, placeAbove ? spaceAbove : spaceBelow),
  );
  panel.style.setProperty(
    "--visualtex-suggestion-max-height",
    \`\${availableHeight}px\`,
  );

  if (placeAbove) {
    panel.classList.add("ML__popover--reverse-direction");
    panel.classList.remove("top-tip");
    panel.classList.add("bottom-tip");
    panel.style.top = \`\${Math.max(
      viewportPadding,
      position.y - position.height - panel.offsetHeight - tipGap,
    )}px\`;
  } else {
    panel.classList.remove("ML__popover--reverse-direction");
    panel.classList.add("top-tip");
    panel.classList.remove("bottom-tip");
    panel.style.top = \`\${Math.max(
      viewportPadding,
      Math.min(position.y + tipGap, viewportBottom - panel.offsetHeight),
    )}px\`;
  }`);
  replace(`  { nextSuggestion, previousSuggestion },`, `  { nextSuggestion, previousSuggestion, visualTexDismissSuggestions: mf => { hideSuggestionPopover(mf); return false; } },`);
  return source;
}
