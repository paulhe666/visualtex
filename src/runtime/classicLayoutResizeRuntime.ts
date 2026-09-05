import { useEditorStore } from "../stores/editorStore";

const TILE_HANDLE_CLASS = "visualtex-classic-tile-resize-handle";
const DOCK_HANDLE_CLASS = "visualtex-classic-dock-resize-handle";
const MIN_TILE_WIDTH = 220;
const MAX_TILE_WIDTH = 720;
const MIN_DOCK_HEIGHT = 140;
const MAX_DOCK_HEIGHT = 560;

let installed = false;
let observer: MutationObserver | null = null;
let refreshFrame = 0;

function clamp(value: number, min: number, max: number) {
  return Math.max(min, Math.min(max, value));
}

function queryFirst(root: ParentNode, selectors: readonly string[]) {
  for (const selector of selectors) {
    const element = root.querySelector<HTMLElement>(selector);
    if (element) return element;
  }
  return null;
}

const TILE_SELECTORS = [
  "[data-classic-tile-panel]",
  "[data-workspace-tile-panel]",
  ".classic-tile-panel",
  ".workspace-tile-panel",
  ".formula-tiles-pane",
  ".tile-panel",
  ".editor-sidebar",
] as const;

const DOCK_SELECTORS = [
  "[data-classic-bottom-dock]",
  "[data-workspace-bottom-dock]",
  ".classic-bottom-dock",
  ".workspace-bottom-dock",
  ".formula-toolbar-dock",
  ".formula-toolbar-shell",
] as const;

function isClassicLayout() {
  const state = useEditorStore.getState();
  return state.editorLayout === "classic";
}

function makeHandle(className: string, ariaLabel: string) {
  const handle = document.createElement("div");
  handle.className = className;
  handle.setAttribute("role", "separator");
  handle.setAttribute("aria-label", ariaLabel);
  handle.tabIndex = 0;
  return handle;
}

function installTileHandle(panel: HTMLElement) {
  if (panel.querySelector(`:scope > .${TILE_HANDLE_CLASS}`)) return;
  panel.dataset.visualtexClassicResizable = "tile";
  const handle = makeHandle(TILE_HANDLE_CLASS, "Resize tile panel");
  panel.append(handle);

  const begin = (clientX: number, pointerId?: number) => {
    const startWidth = panel.getBoundingClientRect().width;
    const startX = clientX;
    const move = (event: PointerEvent) => {
      const next = clamp(
        Math.round(startWidth + event.clientX - startX),
        MIN_TILE_WIDTH,
        MAX_TILE_WIDTH,
      );
      useEditorStore.getState().setClassicTileWidth(next);
    };
    const end = () => {
      window.removeEventListener("pointermove", move, true);
      window.removeEventListener("pointerup", end, true);
      window.removeEventListener("pointercancel", end, true);
      if (pointerId !== undefined && handle.hasPointerCapture?.(pointerId)) {
        handle.releasePointerCapture(pointerId);
      }
    };
    window.addEventListener("pointermove", move, true);
    window.addEventListener("pointerup", end, true);
    window.addEventListener("pointercancel", end, true);
  };

  handle.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    handle.setPointerCapture?.(event.pointerId);
    begin(event.clientX, event.pointerId);
  });
}

function installDockHandle(panel: HTMLElement) {
  if (panel.querySelector(`:scope > .${DOCK_HANDLE_CLASS}`)) return;
  panel.dataset.visualtexClassicResizable = "dock";
  const handle = makeHandle(DOCK_HANDLE_CLASS, "Resize bottom toolbar");
  panel.prepend(handle);

  const begin = (clientY: number, pointerId?: number) => {
    const startHeight = panel.getBoundingClientRect().height;
    const startY = clientY;
    const move = (event: PointerEvent) => {
      const next = clamp(
        Math.round(startHeight + startY - event.clientY),
        MIN_DOCK_HEIGHT,
        MAX_DOCK_HEIGHT,
      );
      useEditorStore.getState().setClassicDockHeight(next);
    };
    const end = () => {
      window.removeEventListener("pointermove", move, true);
      window.removeEventListener("pointerup", end, true);
      window.removeEventListener("pointercancel", end, true);
      if (pointerId !== undefined && handle.hasPointerCapture?.(pointerId)) {
        handle.releasePointerCapture(pointerId);
      }
    };
    window.addEventListener("pointermove", move, true);
    window.addEventListener("pointerup", end, true);
    window.addEventListener("pointercancel", end, true);
  };

  handle.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    handle.setPointerCapture?.(event.pointerId);
    begin(event.clientY, event.pointerId);
  });
}

function refreshClassicLayoutHandles() {
  refreshFrame = 0;
  if (!isClassicLayout()) return;
  const workspace = queryFirst(document, [
    "[data-editor-workspace]",
    ".editor-workspace",
    ".workspace-shell",
    "main",
  ]);
  if (!workspace) return;
  const tilePanel = queryFirst(workspace, TILE_SELECTORS);
  const dockPanel = queryFirst(workspace, DOCK_SELECTORS);
  if (tilePanel) installTileHandle(tilePanel);
  if (dockPanel && dockPanel !== tilePanel) installDockHandle(dockPanel);
}

function scheduleRefresh() {
  if (refreshFrame) return;
  refreshFrame = window.requestAnimationFrame(refreshClassicLayoutHandles);
}

export function installClassicLayoutResizeRuntime() {
  if (installed || typeof document === "undefined") return;
  installed = true;
  scheduleRefresh();
  observer = new MutationObserver(scheduleRefresh);
  observer.observe(document.documentElement, { childList: true, subtree: true });
  useEditorStore.subscribe((state, previous) => {
    if (state.editorLayout !== previous.editorLayout) scheduleRefresh();
  });
}

installClassicLayoutResizeRuntime();
