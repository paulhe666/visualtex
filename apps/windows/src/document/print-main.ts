// The PDF's print page (print.html, a hidden window): it receives the print copy of the page view
// (printCopy.ts), waits for its fonts and images, and says when it is ready for the platform's printing.
import "../editor/visualTexLayoutRuntime";
import "../styles.css";
import "../../../shared/visualtex-app/src/document/document.css";
import "../../../shared/tex-layout/data/profiles/ntx/fonts/fonts.css";
import "../../../shared/visualtex-app/src/document/print.css";
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
  synthesizeCjkFaces(pages);
  await Promise.all([...pages.querySelectorAll("img")].map((img) => img.decode().catch(() => undefined)));
  await document.fonts.ready;
  // fonts that start loading only once text in them is laid out
  pages.getBoundingClientRect();
  await document.fonts.ready;
  return pages.querySelectorAll(".vtx-print-page").length;
}

// Chromium's PDF printing drops synthesized faces: bold and italic Chinese in a font with no such face (SimSun, KaiTi,
// ...) come out regular, though the page shows them. Such text gets them drawn: a stroke for bold (the glyphs keep
// their advance, so no line moves), a slant on each character for italic.
const CJK = /[\u2e80-\u303f\u3040-\u9fff\uac00-\ud7af\uf900-\ufaff\ufe30-\ufe4f\uff00-\uffef]|[\ud840-\ud87f][\udc00-\udfff]/;
const CJK_RUNS = new RegExp(`((?:${CJK.source})+)`, "g");
/** Chinese, Japanese and Korean fonts with their own bold face */
const CJK_BOLD_FACES = /yahei|dengxian|等线|jhenghei|yu gothic|meiryo|malgun|source han|思源|noto (?:sans|serif) (?:cjk|sc|tc|jp|kr)|pingfang|songti|heiti sc|hiragino|微软雅黑/i;
const CJK_FAMILIES = /simsun|nsimsun|宋体|新宋体|simhei|黑体|kaiti|楷体|fangsong|仿宋|yahei|雅黑|dengxian|等线|jhenghei|mingliu|pmingliu|yu gothic|yu mincho|meiryo|ms gothic|ms mincho|malgun|batang|gulim|source han|思源|noto (?:sans|serif) (?:cjk|sc|tc|jp|kr)|pingfang|songti|heiti|kaiti sc|hiragino|st(?:song|kaiti|fangsong|zhongsong|heiti|xihei)|华文/i;
const NOT_TEXT = ".vtx-math-inline, .vtx-math-display, .vtx-note-math, .vtx-code-block, svg, [data-print-hidden]";

function cjkFamily(fontFamily: string): string | null {
  for (const family of fontFamily.split(",")) {
    const name = family.trim().replace(/^["']|["']$/g, "");
    if (CJK_FAMILIES.test(name)) return name;
  }
  return null;
}

function synthesizeCjkFaces(root: HTMLElement) {
  // generated text (the reference list's title)
  for (const element of root.querySelectorAll<HTMLElement>(".vtx-bibliography-title")) {
    if (element.closest(NOT_TEXT)) continue;
    const style = getComputedStyle(element, "::before");
    const family = cjkFamily(style.fontFamily);
    if (CJK.test(style.content) && Number(style.fontWeight) >= 600 && !(family && CJK_BOLD_FACES.test(family))) element.classList.add("vtx-print-bold-before");
  }
  const texts: Text[] = [];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) if (CJK.test(node.nodeValue ?? "")) texts.push(node as Text);
  for (const text of texts) {
    const parent = text.parentElement;
    if (!parent || parent.closest(NOT_TEXT)) continue;
    const style = getComputedStyle(parent);
    const family = cjkFamily(style.fontFamily);
    const bold = Number(style.fontWeight) >= 600 && !(family && CJK_BOLD_FACES.test(family));
    const italic = style.fontStyle !== "normal";
    if (!bold && !italic) continue;
    const pieces = (text.nodeValue ?? "").split(CJK_RUNS);
    const fragment = document.createDocumentFragment();
    pieces.forEach((piece, index) => {
      if (!piece) return;
      if (index % 2 === 0) {
        fragment.append(piece);
        return;
      }
      const run = document.createElement("span");
      if (bold) run.className = "vtx-print-bold";
      if (italic) for (const char of piece) {
        const slanted = document.createElement("span");
        slanted.className = "vtx-print-italic";
        slanted.textContent = char;
        run.append(slanted);
      }
      else run.textContent = piece;
      fragment.append(run);
    });
    text.replaceWith(fragment);
  }
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
