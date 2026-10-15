import { commandRegistry } from "./src/autocomplete/commandRegistry";
import { latestSharedCommands } from "./src/autocomplete/latestSharedCommands";
import { COMMAND_FREQUENCY } from "./src/autocomplete/commandFrequency.generated";

/** Compile discovery, preview and completion into MathLive's own command UI. */
export function patchVisualTexMathLiveSemanticCompletion(source: string) {
  const replace = (from: string, to: string) => {
    if (source.split(from).length !== 2) throw new Error("MathLive semantic completion anchor changed: " + from.slice(0, 70));
    source = source.replace(from, to);
  };
  const nativeCompletionTemplateOverrides = new Map<string, string>([
    [
      "\\begin{cases}",
      "\\begin{cases}\\placeholder{} & \\placeholder{}\\end{cases}",
    ],
  ]);
  // Shared variants (\bigcap with limits, three-row cases, ...) are found by
  // their command but never claim an exactly typed command: \bigcap + Space
  // still inserts the bare operator, the variant follows it.
  const variantIds = new Set(latestSharedCommands.map((c) => c.id));
  const records = commandRegistry.map(c => ({
    exact: !variantIds.has(c.id),
    id: c.id,
    command: c.command,
    latex: nativeCompletionTemplateOverrides.get(c.command) ?? c.insertTemplate,
    preview: c.previewLatex,
    // Search can still use translated keywords, but the native list only
    // presents the command and its mathematical preview.
    label: "", words: [c.labelEn, ...c.aliases, ...c.keywords], priority: c.defaultPriority,
  }));
  const start = source.indexOf("function suggest(mf, s) {");
  const end = source.indexOf("function parseParameterTemplateArgument", start);
  if (start < 0 || end < 0) throw new Error("MathLive suggest anchor changed");
  source = source.slice(0, start) + `const visualTexCompletionCatalog = ${JSON.stringify(records)};\n` +
    `const VISUALTEX_COMMAND_FREQUENCY = ${JSON.stringify(COMMAND_FREQUENCY)};\n` + String.raw`
function visualTexCompletionInfo(mf, latex) {
  return mf.visualTexCompletionRecords?.get(latex) || {
    id: "mathlive-native:" + latex,
    command: latex,
    preview: latex,
    label: "",
  };
}
// Ranking follows Overleaf's LaTeX autocomplete: CodeMirror's fuzzy match
// score decides the tier (exact > prefix > case-folded prefix > word starts >
// substring > scattered), and within a tier the command's corpus frequency
// (boost, scripts/build_command_frequency.mts) decides the order.
//
// visualTexFuzzyScore is a port of FuzzyMatcher.match from
// @codemirror/autocomplete (MIT, (C) Marijn Haverbeke and others), scores
// only. Higher is better; null means no match.
var VISUALTEX_PENALTY = { Gap: -1100, NotStart: -700, CaseFold: -200, ByWord: -100, NotFull: -100 };
function visualTexFuzzyScore(word, pattern) {
  const P = VISUALTEX_PENALTY;
  if (!pattern.length) return P.NotFull;
  if (word.length < pattern.length) return null;
  const chars = [...pattern].map((c) => c.charCodeAt(0));
  const folded = [...pattern].map((c) => {
    const upper = c.toUpperCase();
    return (upper === c ? c.toLowerCase() : upper).charCodeAt(0);
  });
  const len = chars.length;
  if (len === 1) {
    const first = word.charCodeAt(0);
    let score = word.length === 1 ? 0 : P.NotFull;
    if (first === chars[0]) {} else if (first === folded[0]) score += P.CaseFold; else return null;
    return score;
  }
  const direct = word.indexOf(pattern);
  if (direct === 0) return word.length === pattern.length ? 0 : P.NotFull;
  const any = [];
  let anyTo = 0;
  if (direct < 0) {
    for (let i = 0, e = Math.min(word.length, 200); i < e && anyTo < len; i++) {
      const next = word.charCodeAt(i);
      if (next === chars[anyTo] || next === folded[anyTo]) any[anyTo++] = i;
    }
    if (anyTo < len) return null;
  }
  const byWord = [];
  let byWordTo = 0, byWordFolded = false;
  let adjacentTo = 0, adjacentStart = -1, adjacentEnd = -1;
  const hasLower = /[a-z]/.test(word);
  let wordAdjacent = true;
  // 0 non-word, 1 upper, 2 lower/digit
  for (let i = 0, e = Math.min(word.length, 200), prevType = 0; i < e && byWordTo < len; i++) {
    const next = word.charCodeAt(i);
    if (direct < 0 && adjacentTo < len) {
      if (next === chars[adjacentTo] || next === folded[adjacentTo]) {
        if (adjacentTo === 0) adjacentStart = i;
        adjacentEnd = i + 1;
        adjacentTo++;
      } else {
        adjacentTo = 0;
      }
    }
    const type = (next >= 48 && next <= 57) || (next >= 97 && next <= 122) ? 2 : next >= 65 && next <= 90 ? 1 : 0;
    if (!i || (type === 1 && hasLower) || (prevType === 0 && type !== 0)) {
      if (chars[byWordTo] === next || (folded[byWordTo] === next && (byWordFolded = true))) byWord[byWordTo++] = i;
      else if (byWord.length) wordAdjacent = false;
    }
    prevType = type;
  }
  if (byWordTo === len && byWord[0] === 0 && wordAdjacent)
    return P.ByWord + (byWordFolded ? P.CaseFold : 0) - word.length;
  if (adjacentTo === len && adjacentStart === 0)
    return P.CaseFold - word.length + (adjacentEnd === word.length ? 0 : P.NotFull);
  if (direct > -1) return P.NotStart - word.length;
  if (adjacentTo === len) return P.CaseFold + P.NotStart - word.length;
  if (byWordTo === len)
    return P.ByWord + (byWordFolded ? P.CaseFold : 0) + P.NotStart + (wordAdjacent ? 0 : P.Gap) - word.length;
  return len === 2 ? null : (any[0] ? P.NotStart : 0) + P.CaseFold + P.Gap - word.length;
}
// Personal habit as an extra boost in the same units as the corpus frequency
// (\frac ~12.6, \sum ~1.9): uses decay with a 30-day half-life and the boost
// is capped, so habits reorder commands inside a match tier but never lift a
// weaker match above a better one, and one use does not reshuffle the list.
function visualTexPersonalBoost(preferences, usageId) {
  if (!preferences?.personalize || !usageId) return 0;
  const usage = preferences.usage?.[usageId];
  if (!usage) return 0;
  const now = Date.now();
  const recentUses = Array.isArray(usage.recentUses) ? usage.recentUses : [];
  const decayed = recentUses.reduce(
    (sum, timestamp) => sum + Math.pow(0.5, Math.max(0, now - timestamp) / 2_592_000_000),
    0,
  );
  return Math.min(2, Math.log2(1 + decayed)) + (usage.pinned ? 50 : 0);
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
  const rawQuery = s.slice(1);
  const frequency = (command) => VISUALTEX_COMMAND_FREQUENCY[command] || 0;
  // Same score: primary catalog entry, then the bare MathLive command, then
  // shared variants; variants of one command keep their catalog priority.
  const tieBreak = (kind, priority) => (kind === "primary" ? 0.002 : kind === "variant" ? -0.002 : 0) + (priority || 0) * 1e-6;
  const parameterizedCatalogCommands = new Set(
    visualTexCompletionCatalog
      .filter((record) => record.exact && record.latex.includes("\\placeholder{}"))
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
    // Raw LaTeX command completion must only consider actual control
    // sequences. Plain-symbol catalog entries such as the physics constant
    // "c" must never outrank the real command "\\c" just because stripping
    // a leading backslash makes their names look identical.
    if (!record.command.startsWith("\\")) continue;
    const name = record.command.slice(1);
    // When the user has typed a literal delimiter or an atomic native
    // symbol exactly, semantic templates must not replace that literal source.
    // Prefix discovery remains unchanged for non-exact queries.
    if (
      exactLiteralDelimiter ||
      (exactAtomicNative && record.command === s && record.latex !== s)
    ) continue;
    let match = visualTexFuzzyScore(name, rawQuery);
    // An exactly typed command belongs to its primary entry and the bare
    // command; a shared variant ranks as a prefix match.
    if (match === 0 && !record.exact) match = VISUALTEX_PENALTY.NotFull;
    // English aliases and keywords (\integral -> \int) are a last tier
    // below every match on the command itself, and need three letters.
    if (match === null && query.length >= 3 && record.words.some((word) =>
      word.toLowerCase().split(/[^a-z0-9]+/).some((term) => term.startsWith(query))))
      match = -3000;
    if (match === null) continue;
    add(
      record,
      match + frequency(record.command) + visualTexPersonalBoost(preferences, record.id) +
        tieBreak(record.exact ? "primary" : "variant", record.priority),
    );
  }
  // Native symbols and user macros remain available, with case-insensitive
  // discovery (del -> delta AND Delta) and case-sensitive exact preference.
  for (const command of [...Object.keys(LATEX_COMMANDS), ...Object.keys(MATH_SYMBOLS), ...Object.keys(mf.options.macros).map(x => "\\" + x)]) {
    if (LATEX_COMMANDS[command]?.infix) continue;
    const match = visualTexFuzzyScore(command.slice(1), rawQuery);
    if (match === null) continue;
    // Text-mode or diagnostic commands that draw nothing inside a formula.
    if (command === "\\emph" || command === "\\error" || command === "\\the") continue;
    // Parameterized commands such as \\ket, \\sqrt, \\vec, \\abs and
    // \\norm must never compete with their editable VisualTeX template as a
    // bare command. Large operators are the intentional exception: \\int,
    // \\sum, ... are meaningful by themselves and therefore keep both forms.
    if (
      parameterizedCatalogCommands.has(command) &&
      !bareOperatorCommands.has(command) &&
      !(exactAtomicNative && command === s)
    ) continue;
    const usageId = "mathlive-native:" + command;
    add(
      { id: usageId, command, latex: command, preview: command, label: "" },
      match + frequency(command) + visualTexPersonalBoost(preferences, usageId) + tieBreak("native", 0),
    );
  }
  let ranked = [...results.values()].sort((a, b) => b.score - a.score || a.command.length - b.command.length || a.command.localeCompare(b.command));
  // Keep case-insensitive discovery for the first two letters (\be), but once
  // the query is refined (\bet) follow the literal typed prefix so a stale
  // sibling such as \Beta drops out. The panel shows exactly this list, so
  // the controller index, the arrow keys and the number keys all agree.
  if (s.length >= 4) {
    const literal = ranked.filter((r) => r.command.startsWith(s));
    if (literal.length) ranked = literal;
  }
  mf.visualTexCompletionRecords = new Map(ranked.map(r => [r.latex, r]));
  return ranked.slice(0, Math.max(6, Math.min(24, preferences?.count || 12))).map(r => r.latex);
}
// Input-method style selection: the candidates are paged by nine and the
// page holding the current candidate is numbered 1-9.
var VISUALTEX_CANDIDATE_PAGE_SIZE = 9;
function visualTexCandidatePageStart(mf) {
  const index = Math.max(0, mf.suggestionIndex || 0);
  return index - (index % VISUALTEX_CANDIDATE_PAGE_SIZE);
}
function visualTexCandidateNumberMarkup(mf, index) {
  const start = visualTexCandidatePageStart(mf);
  const number = index >= start && index < start + VISUALTEX_CANDIDATE_PAGE_SIZE ? String(index - start + 1) : "";
  return '<span class="ML__popover__index" aria-hidden="true">' + number + "</span>";
}
// -1: the key is not a candidate pick. Otherwise the index of the picked
// candidate, or -2 when the digit has no candidate on the current page (the
// key is swallowed, as an input method does).
function visualTexCandidatePanelOpen(mf) {
  const panel = document.getElementById("mathlive-suggestion-popover");
  if (!panel || panel.visualTexMathfield !== mf || panel.getAttribute("aria-hidden") === "true") return false;
  return (mf.visualTexCompletion?.suggestions || []).length > 0;
}
function visualTexCandidateIndexForDigit(mf, evt) {
  if (mf.model.mode !== "latex" || evt.ctrlKey || evt.metaKey || evt.altKey || evt.isComposing) return -1;
  if (!/^[1-9]$/.test(evt.key || "")) return -1;
  if (!visualTexCandidatePanelOpen(mf)) return -1;
  const suggestions = mf.visualTexCompletion.suggestions;
  const index = visualTexCandidatePageStart(mf) + Number(evt.key) - 1;
  return index < suggestions.length ? index : -2;
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
  replace(`"class=ML__popover__current" : ""}><span class="ML__popover__latex">`, `"class=ML__popover__current" : ""}>\${visualTexCandidateNumberMarkup(mf, i)}<span class="ML__popover__latex">`);
  // Number keys pick from the current page, then commit exactly like Space.
  // Tab while typing a command commits the highlighted candidate like Space;
  // in a font scope (\mathbb{...}) with no candidate list it confirms the
  // scope like Enter.
  replace(`  const keystroke = keyboardEventToString(evt);
  if (evt.isTrusted) {
    validateKeyboardLayout(evt);`, `  let keystroke = keyboardEventToString(evt);
  if (keystroke === "[Tab]" && model.mode === "latex")
    keystroke = mathfield.visualTexSourceWrapper && !visualTexCandidatePanelOpen(mathfield) ? "[Enter]" : "[Space]";
  if (evt.isTrusted) {
    validateKeyboardLayout(evt);`);
  replace(`    if (keystroke === "[Space]" && model.mode === "latex" && mathfield.visualTexSourceWrapper) {`, `    const visualTexPickedCandidate = visualTexCandidateIndexForDigit(mathfield, evt);
    if (visualTexPickedCandidate === -2) {
      if (evt.preventDefault) { evt.preventDefault(); evt.stopPropagation(); }
      return false;
    }
    if (visualTexPickedCandidate >= 0) {
      mathfield.suggestionIndex = visualTexPickedCandidate;
      keystroke = "[Space]";
    }
    if (keystroke === "[Space]" && model.mode === "latex" && mathfield.visualTexSourceWrapper) {`);
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
      const insertionLatex =
        item.dataset.visualtexInsertionLatex || item.dataset.command;
      const index = entries.indexOf(insertionLatex);
      if (index < 0) return;
      owner.suggestionIndex = index;
      complete(owner, "accept-all"); owner.dirty = true; owner.focus();
      requestUpdate(owner);
    });
  }
  panel.visualTexMathfield = mf;
  panel.setAttribute("aria-hidden", "false");
  panel.innerHTML = globalThis.MathfieldElement.createHTML(html);
  for (const item of panel.querySelectorAll("li[data-command]")) {
    const insertionLatex = item.dataset.command || "";
    const info = visualTexCompletionInfo(mf, insertionLatex);
    item.dataset.visualtexInsertionLatex = insertionLatex;
    item.dataset.command = info.command || insertionLatex;
  }
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
  // The LaTeX source editor offers the same commands as the visual editor and
  // needs each command's argument count to insert {} slots.
  replace(`export {\n  MathDivElement,`, String.raw`function visualTexLatexCommandSignatures() {
  const hidden = new Set(["\\emph", "\\error", "\\the"]);
  const signatures = [];
  for (const [name, definition] of Object.entries(LATEX_COMMANDS)) {
    if (definition.infix || hidden.has(name)) continue;
    const params = definition.params || [];
    signatures.push([name, params.filter((p) => !p.isOptional).length, params.filter((p) => p.isOptional).length]);
  }
  for (const name of Object.keys(MATH_SYMBOLS)) signatures.push([name, 0, 0]);
  for (const [name, definition] of Object.entries(getMacros())) signatures.push(["\\" + name, definition?.args || 0, 0]);
  return signatures;
}
export {
  visualTexLatexCommandSignatures,
  MathDivElement,`);
  replace(`  { nextSuggestion, previousSuggestion },`, `  { nextSuggestion, previousSuggestion, visualTexDismissSuggestions: mf => { hideSuggestionPopover(mf); return false; } },`);
  return source;
}
