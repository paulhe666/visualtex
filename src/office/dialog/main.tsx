import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "../../styles.css";
import { configureOcrTransport } from "../../ocr/ocrService";
import { officeOcrTransport } from "../api/ocrHttpTransport";
import { OfficeDialogApp } from "./OfficeDialogApp";

configureOcrTransport(officeOcrTransport);

function mount() {
  const root = document.getElementById("root");
  if (!root) throw new Error("Missing Office Dialog root element.");
  createRoot(root).render(
    <StrictMode>
      <OfficeDialogApp />
    </StrictMode>,
  );
}

const isVstoDesktopRuntime =
  new URLSearchParams(window.location.search).get("runtime") === "vsto-desktop";
const officeRuntime = typeof Office === "undefined" ? null : Office;

if (isVstoDesktopRuntime) {
  mount();
} else if (officeRuntime?.onReady) {
  void officeRuntime.onReady().then(mount).catch((error) => {
    const root = document.getElementById("root");
    if (root) {
      root.textContent =
        error instanceof Error
          ? error.message
          : "Unable to initialize the VisualTeX Office Dialog.";
    }
  });
} else {
  mount();
}
