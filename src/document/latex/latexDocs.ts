// A LaTeX document in the library (personal edition): the library keeps a small pointer as the document's text
// — the project folder and its main file; the folder's files are the document. The ids of LaTeX documents are
// also remembered here, so the library can mark them without reading every document.
import { createSafeStorage } from "../../runtime/safeStorage";

export const LATEX_POINTER_FORMAT = "visualtex-latex-project";

export interface LatexPointer {
  format: typeof LATEX_POINTER_FORMAT;
  version: 1;
  /** the project folder (absolute) */
  root: string;
  /** the main file, relative to the folder */
  main: string;
  /** a chosen engine; null: from the document */
  engine: string | null;
  /** \write18 for minted and the like: the user's choice per document */
  shellEscape: boolean;
  /** the files open in tabs */
  tabs?: string[];
}

export function latexPointer(root: string, main: string, extra: Partial<LatexPointer> = {}): LatexPointer {
  return { format: LATEX_POINTER_FORMAT, version: 1, root, main, engine: null, shellEscape: false, ...extra };
}

export function parseLatexPointer(text: string): LatexPointer | null {
  if (!text.includes(LATEX_POINTER_FORMAT)) return null;
  try {
    const value = JSON.parse(text) as Partial<LatexPointer>;
    if (value.format !== LATEX_POINTER_FORMAT || typeof value.root !== "string" || typeof value.main !== "string") return null;
    return latexPointer(value.root, value.main, {
      engine: typeof value.engine === "string" ? value.engine : null,
      shellEscape: value.shellEscape === true,
      tabs: Array.isArray(value.tabs) ? value.tabs.filter((t): t is string => typeof t === "string") : undefined,
    });
  } catch {
    return null;
  }
}

export function latexPointerText(pointer: LatexPointer): string {
  return JSON.stringify(pointer, null, 1);
}

const storage = createSafeStorage();
const KINDS_KEY = "visualtex.document.latex-ids.v1";

function readIds(): Set<string> {
  try {
    return new Set(JSON.parse(storage.getItem(KINDS_KEY) ?? "[]") as string[]);
  } catch {
    return new Set();
  }
}

export function rememberLatexDocument(id: string) {
  const ids = readIds();
  if (ids.has(id)) return;
  ids.add(id);
  storage.setItem(KINDS_KEY, JSON.stringify([...ids]));
}

export function isLatexDocument(id: string): boolean {
  return readIds().has(id);
}

/** The main file of a new document. */
export function newDocumentSource(isEn: boolean): string {
  if (isEn)
    return [
      "\\documentclass[11pt]{article}",
      "\\usepackage{amsmath,amssymb,graphicx}",
      "\\usepackage{hyperref}",
      "",
      "\\title{Untitled}",
      "\\author{}",
      "\\date{\\today}",
      "",
      "\\begin{document}",
      "\\maketitle",
      "",
      "\\section{Introduction}",
      "",
      "",
      "\\end{document}",
      "",
    ].join("\n");
  return [
    "\\documentclass[UTF8,a4paper,12pt]{ctexart}",
    "\\usepackage{amsmath,amssymb,graphicx}",
    "\\usepackage{hyperref}",
    "",
    "\\title{未命名}",
    "\\author{}",
    "\\date{\\today}",
    "",
    "\\begin{document}",
    "\\maketitle",
    "",
    "\\section{引言}",
    "",
    "",
    "\\end{document}",
    "",
  ].join("\n");
}

/** The title a project gets in the library: its \title, else the main file's name. */
export function projectTitle(source: string, main: string): string {
  const m = /\\title\s*(?:\[[^\]]*\])?\s*\{((?:[^{}]|\{[^{}]*\})*)\}/.exec(source.replace(/(^|[^\\])%.*$/gm, "$1"));
  const title = m ? m[1].replace(/\\(?:\\|thanks\{[^}]*\}|[A-Za-z]+\*?)/g, " ").replace(/[{}~]/g, " ").replace(/\s+/g, " ").trim() : "";
  return (title || main.split("/").pop()!.replace(/\.tex$/i, "")).slice(0, 80);
}
