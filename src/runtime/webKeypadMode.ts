import { safeStorage } from "./safeStorage";
import { useEditorStore } from "../stores/editorStore";
import { formatLatexLines } from "../clipboard/LatexCopyService";

const STORAGE_KEY = "visualtex.web-keypad-mode.v1";
export const WEB_KEYPAD_MODE_CHANGED_EVENT = "visualtex-web-keypad-mode-changed";
const CHANNEL = "visualtex-web-keypad-mode";
let installed = false;

export function readWebKeypadMode() {
  return safeStorage.getItem(STORAGE_KEY) === "true";
}

export function writeWebKeypadMode(enabled: boolean) {
  safeStorage.setItem(STORAGE_KEY, enabled ? "true" : "false");
  if (typeof window !== "undefined") {
    window.dispatchEvent(
      new CustomEvent(WEB_KEYPAD_MODE_CHANGED_EVENT, { detail: enabled }),
    );
    if (typeof BroadcastChannel !== "undefined") {
      const channel = new BroadcastChannel(CHANNEL);
      channel.postMessage(enabled);
      channel.close();
    }
  }
}

export function subscribeWebKeypadMode(listener: (enabled: boolean) => void) {
  if (typeof window === "undefined") return () => undefined;
  const local = (event: Event) => {
    listener(
      (event as CustomEvent<boolean>).detail ?? readWebKeypadMode(),
    );
  };
  const storage = (event: StorageEvent) => {
    if (event.key === STORAGE_KEY) listener(readWebKeypadMode());
  };
  window.addEventListener(WEB_KEYPAD_MODE_CHANGED_EVENT, local);
  window.addEventListener("storage", storage);
  const channel =
    typeof BroadcastChannel === "undefined" ? null : new BroadcastChannel(CHANNEL);
  if (channel) channel.onmessage = () => listener(readWebKeypadMode());
  return () => {
    window.removeEventListener(WEB_KEYPAD_MODE_CHANGED_EVENT, local);
    window.removeEventListener("storage", storage);
    channel?.close();
  };
}

function showCopiedToast() {
  if (typeof document === "undefined") return;
  const existing = document.getElementById("visualtex-web-keypad-toast");
  existing?.remove();
  const toast = document.createElement("div");
  toast.id = "visualtex-web-keypad-toast";
  toast.className = "visualtex-web-keypad-toast";
  toast.textContent =
    useEditorStore.getState().language === "en"
      ? "LaTeX copied"
      : "LaTeX 源码已复制";
  toast.setAttribute("role", "status");
  document.body.append(toast);
  window.setTimeout(() => toast.remove(), 1200);
}

async function copyCurrentLatex() {
  const state = useEditorStore.getState();
  const lines = state.lines.map((line) => line.latex);
  const latex = formatLatexLines(lines, state.latexCodeFormat);
  if (!latex.trim()) return;
  if (!navigator.clipboard?.writeText) {
    throw new Error("Clipboard write is unavailable.");
  }
  await navigator.clipboard.writeText(latex);
  showCopiedToast();
}

export function installWebKeypadModeRuntime() {
  if (installed || typeof window === "undefined") return;
  installed = true;
  window.addEventListener(
    "keydown",
    (event) => {
      if (!readWebKeypadMode()) return;
      if (
        event.key.toLowerCase() !== "s" ||
        !(event.ctrlKey || event.metaKey) ||
        event.altKey ||
        event.shiftKey
      ) {
        return;
      }
      event.preventDefault();
      event.stopImmediatePropagation();
      void copyCurrentLatex().catch(() => undefined);
    },
    true,
  );
}

installWebKeypadModeRuntime();
