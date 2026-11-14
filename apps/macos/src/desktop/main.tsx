import { EditorSessionProvider } from "../../../shared/visualtex-app/src/history/EditorSession";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
// formulas are laid out by the shared engine; installed before any math field exists
import "../../../shared/visualtex-app/src/editor/visualTexLayoutRuntime";
import "../../../shared/visualtex-app/src/styles/index.ts";
import { configureOcrTransport } from "../ocr/ocrService";
import { desktopOcrTransport } from "../ocr/ocrTransport";
import { OfficeDialogApp } from "../office/dialog/OfficeDialogApp";
import { OfficeDocumentImportApp } from "../office/documentImport/OfficeDocumentImportApp";
import { WordLatexRedrawApp } from "../office/redraw/WordLatexRedrawApp";
import { DesktopApp } from "./DesktopApp";
import { applyDocumentTheme, readSynchronizedTheme } from "../../../shared/visualtex-app/src/themeSync";
import { VisualTexErrorBoundary } from "../../../shared/visualtex-app/src/runtime/VisualTexErrorBoundary";
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
    <EditorSessionProvider>
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
    </EditorSessionProvider>
  </StrictMode>,
);
