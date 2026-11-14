// The document as PDF (DESIGN.md §12): the print copy of the page view (printCopy.ts) goes to the hidden print
// window (print.html), which says when its fonts and images are in; the app then has the platform's web view print
// that window to the file (Mac: WKWebView's print operation; Windows: WebView2's PrintToPdf).
import { invoke } from "@tauri-apps/api/core";
import { emitTo, listen } from "@tauri-apps/api/event";
import type { EditorView } from "prosemirror-view";
import { closeFormula, collapseTransaction, hideSourceMarkers, openFormulaId } from "../../../visualtex-document/src/index.ts";
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

/** Every formula closed (the copy shows their static markup) and no Markdown shown around the caret (the source
 *  editor's markers all hidden); the pages laid out again after that. */
async function settleForPrint(view: EditorView) {
  const hide = collapseTransaction(view.state);
  if (hide) view.dispatch(hide);
  const open = openFormulaId(view.state);
  if (open) view.dispatch(closeFormula(view.state.tr));
  view.dispatch(hideSourceMarkers(view.state));
  await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
}

/** Print the document's pages to `path`. */
export async function documentToPdf(view: EditorView, path: string, title: string): Promise<number> {
  await settleForPrint(view);
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
