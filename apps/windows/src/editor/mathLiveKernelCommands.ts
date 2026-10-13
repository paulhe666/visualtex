import type { MathfieldElement, Selection, Selector } from "mathlive";
import type { CommandUsage } from "../types/command";

export interface MathLivePersistentTypingStyle {
  bold: boolean;
  // null follows MathLive's normal math alphabet; false forces upright.
  italic: boolean | null;
  color: string | null;
  backgroundColor: string | null;
}

const STABLE_NATIVE_INPUT_POPOVER_ID =
  "visualtex-native-input-suggestion-popover";
const MATHLIVE_SUGGESTION_POPOVER_ID = "mathlive-suggestion-popover";

let stableNativeInputPopoverFrame = 0;
let stableNativeInputPopoverHideTimer = 0;
let nativeInputPopoverBodyObserver: MutationObserver | null = null;
let nativeInputPopoverSourceObserver: MutationObserver | null = null;
let observedNativeInputPopoverSource: HTMLElement | null = null;
let nativeInputPopoverKeydownInstalled = false;

function nativeInputPopoverSource() {
  return document.getElementById(MATHLIVE_SUGGESTION_POPOVER_ID);
}

function nativeInputPopoverInsertionLatex(item: HTMLElement | null) {
  return (
    item?.dataset.visualtexInsertionLatex ??
    item?.dataset.command ??
    ""
  );
}

function nativeInputPopoverCommands(panel: HTMLElement | null) {
  return Array.from(
    panel?.querySelectorAll<HTMLElement>("li[data-command]") ?? [],
  ).map((item) => nativeInputPopoverInsertionLatex(item));
}

type MathLiveCompletionOwner = {
  suggestionIndex?: number;
  visualTexCompletion?: {
    command?: string;
    suggestions?: string[];
  } | null;
  visualTexCompletionRecords?: Map<
    string,
    { command?: string }
  >;
};

function nativeInputPopoverOwner(source: HTMLElement | null) {
  return (
    source as
      | (HTMLElement & { visualTexMathfield?: MathLiveCompletionOwner })
      | null
  )?.visualTexMathfield;
}

// The kernel's suggest() already narrows refined queries to the literal typed
// prefix, so the visible frame shows MathLive's list one-to-one and the
// controller index, arrow keys and number keys address the same rows.
function stableNativeInputPopoverCommands(source: HTMLElement) {
  return Array.from(source.querySelectorAll<HTMLElement>("li[data-command]"));
}

const CANDIDATE_PAGE_SIZE = 9;

// Mirrors the kernel's numbering: the page holding the selected candidate is
// numbered 1-9 (arrow keys can move the selection onto the next page without
// MathLive re-rendering the list).
function numberCandidatePage(panel: HTMLElement, selectedCommand: string) {
  const items = Array.from(panel.querySelectorAll<HTMLElement>("li[data-command]"));
  const selected = Math.max(
    0,
    items.findIndex(
      (item) => nativeInputPopoverInsertionLatex(item) === selectedCommand,
    ),
  );
  const start = selected - (selected % CANDIDATE_PAGE_SIZE);
  items.forEach((item, index) => {
    const badge = item.querySelector<HTMLElement>(".ML__popover__index");
    if (!badge) return;
    const number =
      index >= start && index < start + CANDIDATE_PAGE_SIZE
        ? String(index - start + 1)
        : "";
    if (badge.textContent !== number) badge.textContent = number;
  });
}

function setStableNativeInputPopoverSelection(
  source: HTMLElement,
  stable: HTMLElement,
  command: string,
  updateOwnerIndex = false,
) {
  if (!command) return;
  const sourceItems = Array.from(
    source.querySelectorAll<HTMLElement>("li[data-command]"),
  );
  if (updateOwnerIndex) {
    const sourceIndex = sourceItems.findIndex(
      (item) => nativeInputPopoverInsertionLatex(item) === command,
    );
    if (sourceIndex >= 0) {
      const owner = nativeInputPopoverOwner(source);
      if (owner) owner.suggestionIndex = sourceIndex;
    }
  }
  for (const item of sourceItems) {
    item.classList.toggle(
      "ML__popover__current",
      nativeInputPopoverInsertionLatex(item) === command,
    );
  }
  for (const item of stable.querySelectorAll<HTMLElement>("li[data-command]")) {
    item.classList.toggle(
      "ML__popover__current",
      nativeInputPopoverInsertionLatex(item) === command,
    );
  }
  numberCandidatePage(source, command);
  numberCandidatePage(stable, command);
}

function ensureStableNativeInputPopover() {
  let panel = document.getElementById(STABLE_NATIVE_INPUT_POPOVER_ID);
  if (panel) return panel;

  panel = document.createElement("div");
  panel.id = STABLE_NATIVE_INPUT_POPOVER_ID;
  panel.setAttribute("aria-hidden", "true");
  panel.addEventListener("pointerdown", (event) => event.preventDefault());
  panel.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof Element)) return;
    const item = target.closest<HTMLElement>("li[data-command]");
    const insertionLatex = nativeInputPopoverInsertionLatex(item);
    if (!insertionLatex) return;
    const sourceItem = Array.from(
      nativeInputPopoverSource()?.querySelectorAll<HTMLElement>(
        "li[data-command]",
      ) ?? [],
    ).find(
      (candidate) =>
        nativeInputPopoverInsertionLatex(candidate) === insertionLatex,
    );
    sourceItem?.dispatchEvent(
      new MouseEvent("click", {
        bubbles: true,
        cancelable: true,
        view: window,
      }),
    );
  });
  document.body.append(panel);
  return panel;
}

function syncStableNativeInputPopoverSelection() {
  const source = nativeInputPopoverSource();
  const stable = document.getElementById(STABLE_NATIVE_INPUT_POPOVER_ID);
  if (!source || !stable) return;
  const allowedCommands = new Set(
    stableNativeInputPopoverCommands(source).map((item) =>
      nativeInputPopoverInsertionLatex(item),
    ),
  );
  let selectedCommand = nativeInputPopoverInsertionLatex(
    source.querySelector<HTMLElement>(
      "li.ML__popover__current[data-command]",
    ),
  );
  if (!allowedCommands.has(selectedCommand)) {
    selectedCommand = nativeInputPopoverInsertionLatex(
      stableNativeInputPopoverCommands(source)[0] ?? null,
    );
    if (!selectedCommand) return;
    // The visible projection has intentionally narrowed the query. Keep the
    // controller aligned with the first still-visible candidate so Space/Enter
    // commits exactly what the user sees selected.
    setStableNativeInputPopoverSelection(
      source,
      stable,
      selectedCommand,
      true,
    );
  } else {
    setStableNativeInputPopoverSelection(source, stable, selectedCommand);
  }
  Array.from(stable.querySelectorAll<HTMLElement>("li[data-command]"))
    .find(
      (item) => nativeInputPopoverInsertionLatex(item) === selectedCommand,
    )
    ?.scrollIntoView({ block: "nearest", inline: "nearest" });
}

function bindNativeInputPopoverSource() {
  const source = nativeInputPopoverSource();
  if (source === observedNativeInputPopoverSource) return;
  nativeInputPopoverSourceObserver?.disconnect();
  observedNativeInputPopoverSource = source;
  if (!source) return;
  nativeInputPopoverSourceObserver = new MutationObserver(() =>
    scheduleStableNativeInputPopoverSync(),
  );
  nativeInputPopoverSourceObserver.observe(source, {
    attributes: true,
    attributeFilter: ["class", "style", "aria-hidden"],
    childList: true,
    characterData: true,
    subtree: true,
  });
}

function installNativeInputPopoverKeydownBridge() {
  if (nativeInputPopoverKeydownInstalled) return;
  nativeInputPopoverKeydownInstalled = true;
  document.addEventListener("keydown", (event) => {
    if (event.key !== "ArrowUp" && event.key !== "ArrowDown") return;
    const source = nativeInputPopoverSource();
    const stable = document.getElementById(STABLE_NATIVE_INPUT_POPOVER_ID);
    if (!source || !stable || !stable.classList.contains("is-visible")) return;

    // Own navigation of the visible projection. MathLive may have already
    // advanced its hidden source during target-phase keydown; the stable frame
    // still contains the pre-keydown selection, so use that as the authoritative
    // starting point and then mirror the chosen command back to MathLive's
    // controller index. This also works while the source temporarily lacks
    // is-visible.
    const stableItems = Array.from(
      stable.querySelectorAll<HTMLElement>("li[data-command]"),
    );
    if (!stableItems.length) return;
    const currentIndex = stableItems.findIndex((item) =>
      item.classList.contains("ML__popover__current"),
    );
    const direction = event.key === "ArrowDown" ? 1 : -1;
    const nextIndex =
      currentIndex < 0
        ? direction > 0
          ? 0
          : stableItems.length - 1
        : (currentIndex + direction + stableItems.length) %
          stableItems.length;
    const nextInsertionLatex = nativeInputPopoverInsertionLatex(
      stableItems[nextIndex] ?? null,
    );
    if (!nextInsertionLatex) return;
    setStableNativeInputPopoverSelection(
      source,
      stable,
      nextInsertionLatex,
      true,
    );
    event.preventDefault();
    scheduleStableNativeInputPopoverSync();
  });
}

function ensureNativeInputPopoverObservers() {
  if (!document.body) return;
  if (!nativeInputPopoverBodyObserver) {
    nativeInputPopoverBodyObserver = new MutationObserver(() => {
      bindNativeInputPopoverSource();
      scheduleStableNativeInputPopoverSync();
    });
    nativeInputPopoverBodyObserver.observe(document.body, { childList: true });
  }
  bindNativeInputPopoverSource();
  installNativeInputPopoverKeydownBridge();
}

function syncStableNativeInputPopover() {
  bindNativeInputPopoverSource();
  const source = nativeInputPopoverSource();
  const stable = ensureStableNativeInputPopover();
  const sourceVisible = Boolean(
    source?.classList.contains("is-visible") &&
      source.querySelector("li[data-command]"),
  );

  if (!source || !sourceVisible) {
    window.clearTimeout(stableNativeInputPopoverHideTimer);
    stableNativeInputPopoverHideTimer = window.setTimeout(() => {
      const latest = nativeInputPopoverSource();
      if (
        latest?.classList.contains("is-visible") &&
        latest.querySelector("li[data-command]")
      ) {
        scheduleStableNativeInputPopoverSync();
        return;
      }
      stable.classList.remove("is-visible");
      stable.setAttribute("aria-hidden", "true");
    }, 64);
    // The source can be hidden for one target-phase keydown while MathLive
    // still changes its selection. Keep the visible frame stable meanwhile.
    syncStableNativeInputPopoverSelection();
    return;
  }

  window.clearTimeout(stableNativeInputPopoverHideTimer);
  source.dataset.visualtexInputPopoverSource = "true";

  const desiredSourceItems = stableNativeInputPopoverCommands(source);
  const desiredCommands = desiredSourceItems.map((item) =>
    nativeInputPopoverInsertionLatex(item),
  );
  const stableCommands = nativeInputPopoverCommands(stable);
  const sameCommands =
    desiredCommands.length === stableCommands.length &&
    desiredCommands.every(
      (command, index) => command === stableCommands[index],
    );
  if (!sameCommands) {
    stable.innerHTML = source.innerHTML;
  }
  const desired = new Set(desiredCommands);
  const owner = nativeInputPopoverOwner(source);
  for (const item of stable.querySelectorAll<HTMLElement>(
    "li[data-command]",
  )) {
    const insertionLatex = nativeInputPopoverInsertionLatex(item);
    if (!desired.has(insertionLatex)) {
      item.remove();
      continue;
    }
    item.dataset.visualtexInsertionLatex = insertionLatex;
    item.dataset.command =
      owner?.visualTexCompletionRecords?.get(insertionLatex)?.command ??
      insertionLatex;
  }
  syncStableNativeInputPopoverSelection();

  stable.classList.toggle("top-tip", source.classList.contains("top-tip"));
  stable.classList.toggle("bottom-tip", source.classList.contains("bottom-tip"));
  stable.style.left = source.style.left;
  stable.style.top = source.style.top;
  const maxCandidateHeight = source.style.getPropertyValue(
    "--visualtex-suggestion-max-height",
  );
  if (maxCandidateHeight) {
    stable.style.setProperty(
      "--visualtex-suggestion-max-height",
      maxCandidateHeight,
    );
  } else {
    stable.style.removeProperty("--visualtex-suggestion-max-height");
  }
  stable.classList.add("is-visible");
  stable.setAttribute("aria-hidden", "false");
}

function scheduleStableNativeInputPopoverSync() {
  if (typeof window === "undefined" || typeof document === "undefined") return;
  ensureNativeInputPopoverObservers();
  window.cancelAnimationFrame(stableNativeInputPopoverFrame);
  stableNativeInputPopoverFrame = window.requestAnimationFrame(
    syncStableNativeInputPopover,
  );
}

export function dismissMathLiveSuggestions(field: MathfieldElement) {
  field.executeCommand("visualTexDismissSuggestions" as Selector);
  scheduleStableNativeInputPopoverSync();
}

export function configureMathLiveCompletion(field: MathfieldElement, preferences: {
  usage: Record<string, CommandUsage>;
  personalize: boolean;
  count: number;
}) {
  (field as MathfieldElement & { visualTexCompletionPreferences?: typeof preferences })
    .visualTexCompletionPreferences = preferences;
  scheduleStableNativeInputPopoverSync();
}

// These commands are registered by vite.mathliveEditingKernel.ts inside MathLive.
// Keep the extension cast here; callers cannot pass arbitrary command names.
export function toggleMathLiveSelectionStyle(
  field: MathfieldElement,
  kind: "bold" | "italic",
  selection: Selection,
) {
  return field.executeCommand("visualTexToggleSelectionStyle" as Selector, kind, selection);
}

export function paintMathLiveDetachedScripts(
  field: MathfieldElement,
  selection: Selection,
) {
  return field.executeCommand("visualTexPaintDetachedScripts" as Selector, selection);
}

export function moveMathLiveToEmptySiblingScript(field: MathfieldElement) {
  return field.executeCommand("visualTexMoveToEmptySiblingScript" as Selector);
}

export function insertMathLiveAlignmentPoint(field: MathfieldElement) {
  return field.executeCommand("visualTexInsertAlignmentPoint" as Selector);
}

export function insertMathLiveRowBreak(
  field: MathfieldElement,
  environment: "aligned" | "gathered",
) {
  return field.executeCommand("visualTexInsertRowBreak" as Selector, environment);
}

export function setMathLivePersistentTypingStyle(
  field: MathfieldElement,
  style: MathLivePersistentTypingStyle,
) {
  const next = {
    bold: style.bold,
    italic: style.italic,
    ...(style.color ? { color: style.color } : {}),
    ...(style.backgroundColor ? { backgroundColor: style.backgroundColor } : {}),
  };
  const target = field as unknown as {
    visualTexPersistentTypingStyle?: typeof next;
    _mathfield?: {
      visualTexPersistentTypingStyle?: typeof next;
    };
  };
  // model.mathfield is MathLive's private controller, not the custom element.
  // Mirror the state on both so the insertion kernel sees it while keeping the
  // public element inspectable in browser regressions.
  target.visualTexPersistentTypingStyle = next;
  if (target._mathfield) {
    target._mathfield.visualTexPersistentTypingStyle = next;
  }
}
