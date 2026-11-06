// The document as PDF (DESIGN.md §12): the print copy of the page view (printCopy.ts) goes to the hidden print
// window (print.html), which says when its fonts and images are in; the app then has the platform's web view print
// that window to the file (Mac: WKWebView's print operation; Windows: WebView2's PrintToPdf).
import { invoke } from "@tauri-apps/api/core";
import { emitTo, listen } from "@tauri-apps/api/event";
import type { EditorView } from "prosemirror-view";
import { closeFormula, openFormulaId } from "../../shared/visualtex-document/src/index.ts";
import { buildPrintCopy } from "./printCopy";
import { PAGE_GAP } from "./DocumentEditor";

const PRINT_WINDOW = "pdf-print";

function once<T>(event: string, timeoutMs: number): { promise: Promise<T>; ready: Promise<void> } {
  let stop: (() => void) | null = null;
  let resolveReady!: () => void;
  const ready = new Promise<void>((resolve) => (resolveReady = resolve));
  const promise = new Promise<T>((resolve, reject) => {
    const timer = window.setTimeout(() => {
      stop?.();
      reject(new Error(`${event}: no answer`));
    }, timeoutMs);
    void listen<T>(event, (e) => {
      window.clearTimeout(timer);
      stop?.();
      resolve(e.payload);
    }).then((unlisten) => {
      stop = unlisten;
      resolveReady();
    });
  });
  return { promise, ready };
}

/** Print the document's pages to `path`. */
export async function documentToPdf(view: EditorView, path: string, title: string): Promise<number> {
  // every formula closed: the copy shows their static markup
  if (openFormulaId(view.state)) {
    view.dispatch(closeFormula(view.state.tr));
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
  }
  const copy = buildPrintCopy(view, PAGE_GAP);
  if (!copy) throw new Error("The page view is not shown");
  const loaded = once<void>("pdf-print-loaded", 20000);
  await loaded.ready;
  await invoke("pdf_print_open");
  try {
    await loaded.promise;
    const ready = once<{ pages?: number; error?: string }>("pdf-print-ready", 60000);
    await ready.ready;
    await emitTo(PRINT_WINDOW, "pdf-print-job", { html: copy.html, widthMm: copy.widthMm, heightMm: copy.heightMm, title });
    const answer = await ready.promise;
    if (answer.error) throw new Error(answer.error);
    await invoke("pdf_print_save", { path, widthMm: copy.widthMm, heightMm: copy.heightMm, title });
    return answer.pages ?? copy.pages;
  } finally {
    void invoke("pdf_print_close").catch(() => undefined);
  }
}

/**
 * The browser's own printing (the web build): the same print copy in the print page (print.html) in a hidden frame,
 * then the browser's print dialog, whose "Save as PDF" makes the file.
 */
export async function printDocumentInBrowser(view: EditorView, title: string): Promise<void> {
  if (openFormulaId(view.state)) {
    view.dispatch(closeFormula(view.state.tr));
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
  }
  const copy = buildPrintCopy(view, PAGE_GAP);
  if (!copy) throw new Error("The page view is not shown");
  const frame = document.createElement("iframe");
  frame.setAttribute("aria-hidden", "true");
  frame.style.cssText = "position:fixed;left:-10000px;top:0;width:1px;height:1px;border:0;visibility:hidden";
  const loaded = new Promise<void>((resolve, reject) => {
    frame.addEventListener("load", () => resolve(), { once: true });
    window.setTimeout(() => reject(new Error("The print page did not load")), 20000);
  });
  frame.src = "/print.html";
  document.body.append(frame);
  try {
    await loaded;
    const target = frame.contentWindow as (Window & { __vtxPrint?: (job: object) => Promise<number> }) | null;
    for (let i = 0; i < 100 && !target?.__vtxPrint; i++) await new Promise((resolve) => setTimeout(resolve, 50));
    if (!target?.__vtxPrint) throw new Error("The print page did not start");
    await target.__vtxPrint({ html: copy.html, widthMm: copy.widthMm, heightMm: copy.heightMm, title });
    target.focus();
    target.print();
  } finally {
    // (the dialog is modal while it is open; the frame goes once it has closed)
    window.setTimeout(() => frame.remove(), 1000);
  }
}
