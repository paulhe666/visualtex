// formulas are laid out by the shared engine; installed before any math field exists
import "../../../shared/visualtex-app/src/editor/visualTexLayoutRuntime";
import "../../../shared/visualtex-app/src/styles/index.ts";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "../ocr/windowsQuickOcrRuntime";
import { configureOcrTransport } from "../../../shared/visualtex-app/src/ocr/ocrService";
import { desktopOcrTransport } from "../../../shared/visualtex-app/src/ocr/ocrTransport";
import { DesktopApp } from "./DesktopApp";
import { applyDocumentTheme, readSynchronizedTheme } from "../../../shared/visualtex-app/src/themeSync";
import { VisualTexErrorBoundary } from "../../../shared/visualtex-app/src/runtime/VisualTexErrorBoundary";
import { EditorSessionProvider } from "../../../shared/visualtex-app/src/history/EditorSession";
import { startSettingsSnapshot } from "./settingsSnapshot";

configureOcrTransport(desktopOcrTransport);
startSettingsSnapshot();
applyDocumentTheme(readSynchronizedTheme());

const root = document.getElementById("root");
if (!root) throw new Error("Missing VisualTeX application root element.");

createRoot(root).render(
  <StrictMode>
    <EditorSessionProvider>
      <VisualTexErrorBoundary>
        <DesktopApp />
      </VisualTexErrorBoundary>
    </EditorSessionProvider>
  </StrictMode>,
);
