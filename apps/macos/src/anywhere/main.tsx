// formulas are laid out by the shared engine; installed before any math field exists
import "../../../shared/visualtex-app/src/editor/visualTexLayoutRuntime";
import { Component, Fragment, StrictMode, type ErrorInfo, type ReactNode } from "react";
import { createRoot } from "react-dom/client";
import { EditorSessionProvider } from "../../../shared/visualtex-app/src/history/EditorSession";
import "../../../shared/visualtex-app/src/styles/index.ts";
import "../../../shared/visualtex-app/src/math/customSymbolRendering";
import "../../../shared/visualtex-app/src/anywhere/anywhere.css";
import { AnywhereApp } from "./AnywhereApp";
import { logAnywhere } from "../../../shared/visualtex-app/src/anywhere/anywhereBridge";
import { configureOcrTransport } from "../../../shared/visualtex-app/src/ocr/ocrService";
import { desktopOcrTransport } from "../../../shared/visualtex-app/src/ocr/ocrTransport";

// Screenshot OCR in the formula goes through the desktop app's OCR, as in the
// main window and the Office editor.
configureOcrTransport(desktopOcrTransport);

/** An error for anywhere.log: its message and the top of its stack. */
function describeError(error: unknown) {
  if (!(error instanceof Error)) return String(error);
  const stack = (error.stack ?? "")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .slice(0, 8)
    .join(" | ");
  return `${error.name}: ${error.message}${stack ? ` @ ${stack}` : ""}`;
}

window.addEventListener("error", (event) => {
  void logAnywhere(`error: ${describeError(event.error ?? event.message)}`);
});
window.addEventListener("unhandledrejection", (event) => {
  void logAnywhere(`unhandled rejection: ${describeError(event.reason)}`);
});

const RECOVERY_WINDOW_MS = 60_000;
const MAX_RECOVERIES = 3;

/**
 * A render error would unmount the whole strip, its activation listeners
 * included, and leave an empty panel: log it and mount the strip again (a
 * fresh page if it keeps failing).
 */
class AnywhereErrorBoundary extends Component<
  { children: ReactNode },
  { failed: boolean; mount: number }
> {
  state = { failed: false, mount: 0 };
  private recoveries: number[] = [];

  static getDerivedStateFromError() {
    return { failed: true };
  }

  componentDidCatch(error: unknown, info: ErrorInfo) {
    const where = (info.componentStack ?? "")
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean)
      .slice(0, 5)
      .join(" < ");
    void logAnywhere(`render error: ${describeError(error)} in ${where}`);
    const now = Date.now();
    this.recoveries = this.recoveries.filter((time) => now - time < RECOVERY_WINDOW_MS);
    if (this.recoveries.length >= MAX_RECOVERIES) {
      window.location.reload();
      return;
    }
    this.recoveries.push(now);
    window.setTimeout(() => this.setState((state) => ({ failed: false, mount: state.mount + 1 })), 0);
  }

  render() {
    if (this.state.failed) return null;
    return <Fragment key={this.state.mount}>{this.props.children}</Fragment>;
  }
}

const root = document.getElementById("root");
if (!root) throw new Error("Missing VisualTeX Anywhere root element.");

createRoot(root).render(
  <StrictMode>
    <AnywhereErrorBoundary>
      <EditorSessionProvider>
        <AnywhereApp />
      </EditorSessionProvider>
    </AnywhereErrorBoundary>
  </StrictMode>,
);
