import { Component, StrictMode, type ErrorInfo, type ReactNode } from "react";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "../styles.css";
import { configureOcrTransport } from "../ocr/ocrService";
import { desktopOcrTransport } from "../ocr/ocrTransport";
import { OfficeDialogApp } from "./dialog/OfficeDialogApp";

configureOcrTransport(desktopOcrTransport);

class OfficeFormulaErrorBoundary extends Component<
  { children: ReactNode },
  { error: Error | null }
> {
  state = { error: null as Error | null };

  static getDerivedStateFromError(error: Error) {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.error("VisualTeX Office formula editor crashed", error, info);
  }

  render() {
    if (!this.state.error) return this.props.children;
    return (
      <main className="office-dialog-state is-error" role="alert">
        <strong>VisualTeX Office 公式编辑器加载失败</strong>
        <p>{this.state.error.stack || this.state.error.message}</p>
      </main>
    );
  }
}

const root = document.getElementById("root");
if (!root) throw new Error("Missing VisualTeX native Office editor root element.");

createRoot(root).render(
  <StrictMode>
    <OfficeFormulaErrorBoundary>
      <OfficeDialogApp />
    </OfficeFormulaErrorBoundary>
  </StrictMode>,
);
