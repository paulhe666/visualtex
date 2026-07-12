import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "mathlive/static.css";
import "../styles.css";
import { configureOcrTransport } from "../ocr/ocrService";
import { desktopOcrTransport } from "../ocr/ocrTransport";
import { DesktopApp } from "./DesktopApp";

configureOcrTransport(desktopOcrTransport);

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <DesktopApp />
  </StrictMode>,
);
