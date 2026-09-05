import { OPEN_CUSTOM_SYMBOL_DESIGNER_EVENT } from "../components/GlobalCustomSymbolDesignerHost";
import { useEditorStore } from "../stores/editorStore";

const BUTTON_CLASS = "visualtex-custom-symbol-toolbar-entry";
let installed = false;
let observer: MutationObserver | null = null;
let frame = 0;

const TOOLBAR_SELECTORS = [
  "[data-formula-toolbar]",
  ".formula-toolbar",
  ".formula-toolbar-shell",
  ".formula-tool-grid",
] as const;

function findToolbar() {
  for (const selector of TOOLBAR_SELECTORS) {
    const toolbar = document.querySelector<HTMLElement>(selector);
    if (toolbar) return toolbar;
  }
  return null;
}

function refresh() {
  frame = 0;
  const existing = document.querySelector<HTMLButtonElement>(`.${BUTTON_CLASS}`);
  if (existing) {
    const isEn = useEditorStore.getState().language === "en";
    existing.title = isEn ? "Custom Symbol Designer" : "自定义字符设计器";
    existing.setAttribute("aria-label", existing.title);
    const label = existing.querySelector<HTMLElement>(".visualtex-custom-symbol-toolbar-label");
    if (label) label.textContent = isEn ? "Custom" : "自定义";
    return;
  }
  const toolbar = findToolbar();
  if (!toolbar) return;
  const button = document.createElement("button");
  button.type = "button";
  button.className = BUTTON_CLASS;
  button.dataset.customSymbolDesignerEntry = "true";
  const isEn = useEditorStore.getState().language === "en";
  button.title = isEn ? "Custom Symbol Designer" : "自定义字符设计器";
  button.setAttribute("aria-label", button.title);
  button.innerHTML = `<span aria-hidden="true">✦</span><span class="visualtex-custom-symbol-toolbar-label">${isEn ? "Custom" : "自定义"}</span>`;
  button.addEventListener("click", () => {
    window.dispatchEvent(new Event(OPEN_CUSTOM_SYMBOL_DESIGNER_EVENT));
  });
  toolbar.append(button);
}

function scheduleRefresh() {
  if (frame) return;
  frame = window.requestAnimationFrame(refresh);
}

export function installCustomSymbolToolbarRuntime() {
  if (installed || typeof document === "undefined") return;
  installed = true;
  scheduleRefresh();
  observer = new MutationObserver(scheduleRefresh);
  observer.observe(document.documentElement, { childList: true, subtree: true });
  useEditorStore.subscribe((state, previous) => {
    if (state.language !== previous.language) scheduleRefresh();
  });
}

installCustomSymbolToolbarRuntime();
