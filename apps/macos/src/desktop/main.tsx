import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "../styles.css";
import { configureOcrTransport } from "../ocr/ocrService";
import { desktopOcrTransport } from "../ocr/ocrTransport";
import { OfficeDialogApp } from "../office/dialog/OfficeDialogApp";
import { DesktopApp } from "./DesktopApp";

configureOcrTransport(desktopOcrTransport);

const root = document.getElementById("root");
if (!root) throw new Error("Missing VisualTeX application root element.");

const view = new URLSearchParams(window.location.search).get("view");
const officeFormulaView = view === "office-formula";
if (officeFormulaView) {
  document.body.classList.add("office-dialog-page");
}

createRoot(root).render(
  <StrictMode>
    {officeFormulaView ? <OfficeDialogApp /> : <DesktopApp />}
  </StrictMode>,
);
