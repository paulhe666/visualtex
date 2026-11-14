import "../../../../shared/visualtex-app/src/editor/visualTexLayoutRuntime";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "../../../../shared/visualtex-app/src/styles/index.ts";
import "../../../../shared/visualtex-app/src/math/customSymbolRendering";
import { configureOcrTransport } from "../../ocr/ocrService";
import { officeOcrTransport } from "../api/ocrHttpTransport";
import { OfficeDialogApp } from "./OfficeDialogApp";
import { installFloatingLayerAutoAvoidance } from "../../../../shared/visualtex-app/src/runtime/floatingLayerAutoAvoidance";
import { DocumentImportApp } from "../documentImport/DocumentImportApp";
import {
  applyDocumentTheme,
  normalizeSynchronizedTheme,
  readSynchronizedTheme,
} from "../../../../shared/visualtex-app/src/themeSync";
import {
  normalizeEditorLayout,
  useEditorStore,
} from "../../../../shared/visualtex-app/src/stores/editorStore";
import { VisualTexErrorBoundary } from "../../../../shared/visualtex-app/src/runtime/VisualTexErrorBoundary";
import { EditorSessionProvider } from "../../../../shared/visualtex-app/src/history/EditorSession";

configureOcrTransport(officeOcrTransport);
installFloatingLayerAutoAvoidance();
const injectedTheme = document
  .querySelector<HTMLMetaElement>('meta[name="visualtex-theme"]')
  ?.content;
const initialTheme = injectedTheme
  ? normalizeSynchronizedTheme(injectedTheme)
  : readSynchronizedTheme();
useEditorStore.getState().setTheme(initialTheme);
applyDocumentTheme(initialTheme);
useEditorStore.getState().setEditorLayout(
  normalizeEditorLayout(
    document
      .querySelector<HTMLMetaElement>('meta[name="visualtex-editor-layout"]')
      ?.content,
  ),
);

function mount() {
  const root = document.getElementById("root");
  if (!root) throw new Error("Missing Office Dialog root element.");
  const runtime = new URLSearchParams(window.location.search).get("runtime");
  createRoot(root).render(
    <StrictMode>
      <EditorSessionProvider>
        <VisualTexErrorBoundary
          title="VisualTeX Office 编辑器加载失败"
          description="Office 编辑器加载失败。"
          logContext="VisualTeX Office dialog crashed"
          enableSafeRestart={false}
        >
          {runtime === "vsto-bulk-import" ? (
            <DocumentImportApp />
          ) : (
            <OfficeDialogApp />
          )}
        </VisualTexErrorBoundary>
      </EditorSessionProvider>
    </StrictMode>,
  );
}

// Windows Office integration is now exclusively driven by the native Ribbon
// COM add-ins. The dialog is hosted by the VisualTeX companion service and no
// longer waits for, imports or executes the Office.js runtime.
mount();
