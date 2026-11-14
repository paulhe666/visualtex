import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "../ocr/windowsQuickOcrRuntime";
import { configureOcrTransport } from "../ocr/ocrService";
import { desktopOcrTransport } from "../ocr/ocrTransport";
import { DesktopApp } from "./DesktopApp";
import { applyDocumentTheme, readSynchronizedTheme } from "../../../shared/visualtex-app/src/themeSync";
import { VisualTexErrorBoundary } from "../../../shared/visualtex-app/src/runtime/VisualTexErrorBoundary";
import { EditorSessionProvider } from "../../../shared/visualtex-app/src/history/EditorSession";

configureOcrTransport(desktopOcrTransport);
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
