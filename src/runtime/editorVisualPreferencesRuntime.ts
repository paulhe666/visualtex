import { useEditorStore } from "../stores/editorStore";

let installed = false;
let unsubscribe: (() => void) | null = null;

function applyEditorVisualPreferences() {
  if (typeof document === "undefined") return;
  const state = useEditorStore.getState();
  const root = document.documentElement;
  root.dataset.visualtexShowLineNumbers = state.showLineNumbers ? "true" : "false";
  root.dataset.visualtexHighlightActiveLine = state.highlightActiveLine
    ? "true"
    : "false";
  root.dataset.visualtexEditorLayout = state.editorLayout;
  root.style.setProperty(
    "--visualtex-formula-inset-left",
    `${state.formulaInsetLeft}px`,
  );
  root.style.setProperty(
    "--visualtex-formula-inset-right",
    `${state.formulaInsetRight}px`,
  );
  root.style.setProperty(
    "--visualtex-formula-row-vertical-inset",
    `${state.formulaRowVerticalInset}px`,
  );
  root.style.setProperty(
    "--visualtex-formula-tool-button-size",
    `${state.formulaToolButtonSize}px`,
  );
  root.style.setProperty(
    "--visualtex-formula-tool-button-padding",
    `${state.formulaToolButtonPadding}px`,
  );
  root.style.setProperty(
    "--visualtex-classic-tile-width",
    `${state.classicTileWidth}px`,
  );
  root.style.setProperty(
    "--visualtex-classic-dock-height",
    `${state.classicDockHeight}px`,
  );
}

export function installEditorVisualPreferencesRuntime() {
  if (installed || typeof document === "undefined") return;
  installed = true;
  applyEditorVisualPreferences();
  unsubscribe = useEditorStore.subscribe((state, previous) => {
    if (
      state.showLineNumbers !== previous.showLineNumbers ||
      state.highlightActiveLine !== previous.highlightActiveLine ||
      state.editorLayout !== previous.editorLayout ||
      state.formulaInsetLeft !== previous.formulaInsetLeft ||
      state.formulaInsetRight !== previous.formulaInsetRight ||
      state.formulaRowVerticalInset !== previous.formulaRowVerticalInset ||
      state.formulaToolButtonSize !== previous.formulaToolButtonSize ||
      state.formulaToolButtonPadding !== previous.formulaToolButtonPadding ||
      state.classicTileWidth !== previous.classicTileWidth ||
      state.classicDockHeight !== previous.classicDockHeight
    ) {
      applyEditorVisualPreferences();
    }
  });
}

export function disposeEditorVisualPreferencesRuntime() {
  unsubscribe?.();
  unsubscribe = null;
  installed = false;
}

installEditorVisualPreferencesRuntime();
