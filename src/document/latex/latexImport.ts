// Opening LaTeX sources as LaTeX documents (personal edition): a .tex file (its folder becomes the project), a
// folder, or a .zip (unpacked into a project folder of the app's). Nothing is converted: the files are the
// document as they are.
import { unzipSync } from "fflate";
import { decodeText } from "../../../shared/visualtex-document/src/index.ts";
import { latexPointer, newDocumentSource, projectTitle, type LatexPointer } from "./latexDocs";
import { bytesToBase64, texService, type ProjectInfo, type TexService } from "./texService";

export interface MadeProject {
  pointer: LatexPointer;
  title: string;
}

function service(): TexService {
  const s = texService();
  if (!s) throw new Error("No TeX in this build");
  return s;
}

const ZIP_SKIP = /(^|\/)(__MACOSX\/|\._|\.DS_Store$)/;

/** A .tex/.ltx file, or a .zip holding one. */
export async function isLatexImport(file: File): Promise<boolean> {
  if (/\.(tex|ltx)$/i.test(file.name)) return true;
  if (!/\.zip$/i.test(file.name)) return false;
  let found = false;
  try {
    unzipSync(new Uint8Array(await file.arrayBuffer()), {
      filter: (entry) => {
        if (/\.(tex|ltx)$/i.test(entry.name) && !ZIP_SKIP.test(entry.name)) found = true;
        return false;
      },
    });
  } catch {
    return false;
  }
  return found;
}

async function made(info: ProjectInfo, main?: string | null): Promise<MadeProject> {
  const s = service();
  const mainPath = main ?? info.main;
  if (!mainPath) throw new Error("No main .tex file (one with \\documentclass and \\begin{document})");
  let title = "";
  try {
    const { base64ToBytes } = await import("./texService");
    title = projectTitle(decodeText(base64ToBytes(await s.read(info.root, mainPath))), mainPath);
  } catch {
    title = mainPath.replace(/\.tex$/i, "");
  }
  return { pointer: latexPointer(info.root, mainPath), title };
}

/** Files (path → bytes) into a new project folder of the app's; the common top folder of an archive dropped. */
async function writeProject(name: string, files: Array<{ path: string; bytes: Uint8Array }>): Promise<ProjectInfo> {
  const s = service();
  const kept = files.filter((f) => !ZIP_SKIP.test(f.path) && !f.path.endsWith("/"));
  const tops = new Set(kept.map((f) => (f.path.includes("/") ? f.path.split("/")[0] : "")));
  const strip = tops.size === 1 && !tops.has("") ? [...tops][0].length + 1 : 0;
  const project = await s.create(name);
  for (const f of kept) await s.write(project.root, f.path.slice(strip), { base64: bytesToBase64(f.bytes) });
  return s.open(project.root);
}

/** A .tex (from its folder when the app knows it) or a .zip, as a project. */
export async function openLatexImport(file: File, folder: string | null): Promise<MadeProject> {
  const s = service();
  const stem = file.name.replace(/\.[^.]+$/, "");
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (/\.zip$/i.test(file.name)) {
    const entries = unzipSync(bytes);
    return made(await writeProject(stem, Object.entries(entries).map(([path, data]) => ({ path, bytes: data }))));
  }
  if (folder) {
    const info = await s.open(folder.replace(/[\\/]$/, "") + (folder.includes("\\") ? "\\" : "/") + file.name, file.name);
    return made(info, file.name);
  }
  // a browser: the file alone
  const info = await writeProject(stem, [{ path: file.name, bytes }]);
  return made(info, file.name);
}

export async function newLatexDocument(isEn: boolean): Promise<MadeProject> {
  const s = service();
  const project = await s.create("latex");
  await s.write(project.root, "main.tex", { text: newDocumentSource(isEn) });
  const info = await s.open(project.root, "main.tex");
  return { pointer: latexPointer(info.root, "main.tex"), title: "" };
}

/** A folder from the system's dialog (the app). */
export async function openLatexFolder(isEn: boolean): Promise<MadeProject | null> {
  const { open } = await import("@tauri-apps/plugin-dialog");
  const chosen = await open({ directory: true, multiple: false, title: isEn ? "Open folder" : "打开文件夹" });
  if (!chosen || Array.isArray(chosen)) return null;
  return made(await service().open(chosen));
}

/** A folder picked in a browser (its files copied into a project folder). */
export function pickLatexFolderInBrowser(): Promise<MadeProject | null> {
  return new Promise((resolve, reject) => {
    const input = document.createElement("input");
    input.type = "file";
    input.multiple = true;
    input.setAttribute("webkitdirectory", "");
    input.addEventListener("change", () => {
      const list = [...(input.files ?? [])];
      if (!list.length) return resolve(null);
      void (async () => {
        const files = await Promise.all(list.map(async (f) => ({ path: f.webkitRelativePath || f.name, bytes: new Uint8Array(await f.arrayBuffer()) })));
        const name = (list[0].webkitRelativePath || "latex").split("/")[0];
        return made(await writeProject(name, files));
      })().then(resolve, reject);
    });
    input.click();
  });
}
