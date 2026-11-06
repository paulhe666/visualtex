// The PDF's print page (print.html, a hidden window): it receives the print copy of the page view
// (printCopy.ts), waits for its fonts and images, and says when it is ready for the platform's printing.
import "../editor/visualTexLayoutRuntime";
import "../styles.css";
import "./document.css";
import "../../shared/tex-layout/data/profiles/ntx/fonts/fonts.css";
import "./print.css";
import { isTauri } from "@tauri-apps/api/core";
import { emit, listen } from "@tauri-apps/api/event";

export interface PrintJob {
  html: string;
  widthMm: number;
  heightMm: number;
  title: string;
}

/** The pages in place, their fonts and images loaded (the page count once laid out). */
export async function renderPrintJob(job: PrintJob): Promise<number> {
  document.title = job.title || "VisualTeX";
  let page = document.getElementById("vtx-print-page-size") as HTMLStyleElement | null;
  if (!page) {
    page = document.createElement("style");
    page.id = "vtx-print-page-size";
    document.head.append(page);
  }
  page.textContent = `@page { size: ${job.widthMm}mm ${job.heightMm}mm; margin: 0; }`;
  const pages = document.getElementById("pages")!;
  pages.innerHTML = job.html;
  await Promise.all([...pages.querySelectorAll("img")].map((img) => img.decode().catch(() => undefined)));
  await document.fonts.ready;
  // fonts that start loading only once text in them is laid out
  pages.getBoundingClientRect();
  await document.fonts.ready;
  return pages.querySelectorAll(".vtx-print-page").length;
}

(window as unknown as { __vtxPrint?: typeof renderPrintJob }).__vtxPrint = renderPrintJob;

if (isTauri()) {
  void listen<PrintJob>("pdf-print-job", async (event) => {
    try {
      const pages = await renderPrintJob(event.payload);
      await emit("pdf-print-ready", { pages });
    } catch (error) {
      await emit("pdf-print-ready", { error: String(error) });
    }
  }).then(() => emit("pdf-print-loaded"));
}
