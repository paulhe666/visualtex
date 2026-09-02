import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "../../styles.css";
import "../../styles-macos-main.css";
import "../../styles-editor-parity.css";
import "../../styles-latest-macos-ui.css";
import "../../styles-windows-shared-latest.css";
import "../../styles-custom-symbol-designer.css";
import "../../math/customSymbolRendering";
import { configureOcrTransport } from "../../ocr/ocrService";
import { officeOcrTransport } from "../api/ocrHttpTransport";
import { OfficeDialogApp } from "./OfficeDialogApp";
import { installFloatingLayerAutoAvoidance } from "../../runtime/floatingLayerAutoAvoidance";
import { DocumentImportApp } from "../documentImport/DocumentImportApp";
import {
  applyDocumentTheme,
  normalizeSynchronizedTheme,
  readSynchronizedTheme,
} from "../../themeSync";
import {
  normalizeEditorLayout,
  useEditorStore,
} from "../../stores/editorStore";
import { VisualTexErrorBoundary } from "../../runtime/VisualTexErrorBoundary";

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
      <VisualTexErrorBoundary
        title="VisualTeX Office 编辑器加载失败"
        description="Office 编辑器在加载界面时遇到了异常。重新打开后若仍出现此页面，请把下面的错误信息发送给开发者。"
        logContext="VisualTeX Office dialog crashed"
      >
        {runtime === "vsto-bulk-import" ? (
          <DocumentImportApp />
        ) : (
          <OfficeDialogApp />
        )}
      </VisualTexErrorBoundary>
    </StrictMode>,
  );
}

// Windows Office integration is now exclusively driven by the native Ribbon
// COM add-ins. The dialog is hosted by the VisualTeX companion service and no
// longer waits for, imports or executes the Office.js runtime.
mount();
