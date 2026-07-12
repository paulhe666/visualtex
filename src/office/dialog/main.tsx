import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "../../styles.css";
import { OfficeDialogApp } from "./OfficeDialogApp";

function mount() {
  const root = document.getElementById("root");
  if (!root) throw new Error("Missing Office Dialog root element.");
  createRoot(root).render(
    <StrictMode>
      <OfficeDialogApp />
    </StrictMode>,
  );
}

void Office.onReady().then(mount).catch((error) => {
  const root = document.getElementById("root");
  if (root) {
    root.textContent =
      error instanceof Error
        ? error.message
        : "Unable to initialize the VisualTeX Office Dialog.";
  }
});
