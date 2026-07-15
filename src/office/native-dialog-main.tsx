import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "../styles.css";
import { configureOcrTransport } from "../ocr/ocrService";
import { desktopOcrTransport } from "../ocr/ocrTransport";
import { OfficeDialogApp } from "./dialog/OfficeDialogApp";

configureOcrTransport(desktopOcrTransport);

const root = document.getElementById("root");
if (!root) throw new Error("Missing VisualTeX native Office editor root element.");

createRoot(root).render(
  <StrictMode>
    <OfficeDialogApp />
  </StrictMode>,
);
