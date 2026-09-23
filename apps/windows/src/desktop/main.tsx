import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "../ocr/windowsQuickOcrRuntime";
import { configureOcrTransport } from "../ocr/ocrService";
import { desktopOcrTransport } from "../ocr/ocrTransport";
import { DesktopApp } from "./DesktopApp";
import { applyDocumentTheme, readSynchronizedTheme } from "../themeSync";
import { VisualTexErrorBoundary } from "../runtime/VisualTexErrorBoundary";

configureOcrTransport(desktopOcrTransport);
applyDocumentTheme(readSynchronizedTheme());

const root = document.getElementById("root");
if (!root) throw new Error("Missing VisualTeX application root element.");

createRoot(root).render(
  <StrictMode>
    <VisualTexErrorBoundary>
      <DesktopApp />
    </VisualTexErrorBoundary>
  </StrictMode>,
);
