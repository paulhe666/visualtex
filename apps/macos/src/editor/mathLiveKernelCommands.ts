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

function nativeInputPopoverCommands(panel: HTMLElement | null) {
  return Array.from(
    panel?.querySelectorAll<HTMLElement>("li[data-command]") ?? [],
  ).map((item) => item.dataset.command ?? "");
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

function stableNativeInputPopoverCommands(source: HTMLElement) {
  const sourceItems = Array.from(
    source.querySelectorAll<HTMLElement>("li[data-command]"),
  );
  const owner = nativeInputPopoverOwner(source);
  const query = owner?.visualTexCompletion?.command ?? "";
  // data-command stores the insertion LaTeX. For compatibility aliases such
  // as \Bbb, that insertion is canonicalized to \mathbb{...}; filter using
  // the completion record's typed command instead or the exact alias vanishes
  // from its own candidate list.
  const typedCommand = (item: HTMLElement) => {
    const latex = item.dataset.command ?? "";
    return owner?.visualTexCompletionRecords?.get(latex)?.command ?? latex;
  };
  // Keep MathLive's broad case-insensitive discovery for the first two typed
  // letters (e.g. \be), but once the user refines the query (\bet) the
  // visible candidate frame follows the literal typed prefix. This prevents a
  // stale uppercase sibling such as \Beta from surviving a lowercase refined
  // query while leaving MathLive's underlying catalog untouched.
  const filtered =
    query.length >= 4
      ? sourceItems.filter((item) => typedCommand(item).startsWith(query))
      : sourceItems;
  return filtered.length ? filtered : sourceItems;
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
      (item) => item.dataset.command === command,
    );
    if (sourceIndex >= 0) {
      const owner = nativeInputPopoverOwner(source);
      if (owner) owner.suggestionIndex = sourceIndex;
    }
  }
  for (const item of sourceItems) {
    item.classList.toggle(
      "ML__popover__current",
      item.dataset.command === command,
    );
  }
  for (const item of stable.querySelectorAll<HTMLElement>("li[data-command]")) {
    item.classList.toggle(
      "ML__popover__current",
      item.dataset.command === command,
    );
  }
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
    const command = item?.dataset.command ?? "";
    if (!command) return;
    const sourceItem = Array.from(
      nativeInputPopoverSource()?.querySelectorAll<HTMLElement>(
        "li[data-command]",
      ) ?? [],
    ).find((candidate) => candidate.dataset.command === command);
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
    stableNativeInputPopoverCommands(source).map(
      (item) => item.dataset.command ?? "",
    ),
  );
  let selectedCommand =
    source.querySelector<HTMLElement>(
      "li.ML__popover__current[data-command]",
    )?.dataset.command ?? "";
  if (!allowedCommands.has(selectedCommand)) {
    selectedCommand =
      stableNativeInputPopoverCommands(source)[0]?.dataset.command ?? "";
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
    .find((item) => item.dataset.command === selectedCommand)
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
    const nextCommand = stableItems[nextIndex]?.dataset.command ?? "";
    if (!nextCommand) return;
    setStableNativeInputPopoverSelection(
      source,
      stable,
      nextCommand,
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
  const desiredCommands = desiredSourceItems.map(
    (item) => item.dataset.command ?? "",
  );
  const stableCommands = nativeInputPopoverCommands(stable);
  const sameCommands =
    desiredCommands.length === stableCommands.length &&
    desiredCommands.every(
      (command, index) => command === stableCommands[index],
    );
  if (!sameCommands) {
    stable.innerHTML = source.innerHTML;
    const desired = new Set(desiredCommands);
    for (const item of stable.querySelectorAll<HTMLElement>(
      "li[data-command]",
    )) {
      if (!desired.has(item.dataset.command ?? "")) item.remove();
    }
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
