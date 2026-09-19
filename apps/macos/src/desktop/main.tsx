import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "../styles.css";
import "../styles-editor-parity.css";
import "../styles-windows-shared-latest.css";
import "../styles-macos-platform-overrides.css";
import { configureOcrTransport } from "../ocr/ocrService";
import { desktopOcrTransport } from "../ocr/ocrTransport";
import { OfficeDialogApp } from "../office/dialog/OfficeDialogApp";
import { OfficeDocumentImportApp } from "../office/documentImport/OfficeDocumentImportApp";
import { WordLatexRedrawApp } from "../office/redraw/WordLatexRedrawApp";
import { DesktopApp } from "./DesktopApp";
import { applyDocumentTheme, readSynchronizedTheme } from "../themeSync";
import { VisualTexErrorBoundary } from "../runtime/VisualTexErrorBoundary";
import { QuickOcrHud } from "../ocr/QuickOcrHud";

configureOcrTransport(desktopOcrTransport);

const root = document.getElementById("root");
if (!root) throw new Error("Missing VisualTeX application root element.");

const view = new URLSearchParams(window.location.search).get("view");
const officeFormulaView = view === "office-formula";
const officeDocumentImportView = view === "office-document-import";
const officeWordLatexRedrawView = view === "office-word-latex-redraw";
const quickOcrHudView = view === "quick-ocr-hud";
if (officeFormulaView || officeDocumentImportView || officeWordLatexRedrawView) {
  document.body.classList.add("office-dialog-page");
  applyDocumentTheme(readSynchronizedTheme());
}
if (quickOcrHudView) {
  document.body.classList.add("quick-ocr-hud-page");
  applyDocumentTheme(readSynchronizedTheme());
}

createRoot(root).render(
  <StrictMode>
    <VisualTexErrorBoundary>
      {quickOcrHudView ? (
        <QuickOcrHud />
      ) : officeFormulaView ? (
        <OfficeDialogApp />
      ) : officeDocumentImportView ? (
        <OfficeDocumentImportApp />
      ) : officeWordLatexRedrawView ? (
        <WordLatexRedrawApp />
      ) : (
        <DesktopApp />
      )}
    </VisualTexErrorBoundary>
  </StrictMode>,
);
