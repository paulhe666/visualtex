let installed = false;

function dispatchQuickOcr() {
  window.dispatchEvent(new CustomEvent("visualtex-quick-ocr"));
}

/**
 * Windows quick OCR remains an interactive React workflow. Silent OCR is
 * deliberately absent from this bridge: Ctrl+Alt+O is registered and executed
 * completely in the native runtime so it keeps working while the main WebView
 * is minimized, hidden, unfocused or not created at all.
 */
export function installWindowsQuickOcrRuntime() {
  if (installed || typeof window === "undefined" || typeof document === "undefined") {
    return;
  }
  installed = true;
  document.addEventListener(
    "keydown",
    (event) => {
      if (event.repeat || event.isComposing) return;
      if (
        event.ctrlKey &&
        event.shiftKey &&
        !event.altKey &&
        !event.metaKey &&
        event.code === "KeyO"
      ) {
        event.preventDefault();
        dispatchQuickOcr();
      }
    },
    true,
  );
}

installWindowsQuickOcrRuntime();
