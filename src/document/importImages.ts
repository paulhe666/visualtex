// The pictures a Markdown or LaTeX file names, found before it is read (images.ts resolves them synchronously while
// the text is read): from the web, from the folder the file was opened from, from its archive. Formats the document
// does not keep (GIF, WebP, SVG…) are drawn into PNG.
import { registerImageBytes, registerImageFile } from "../../shared/visualtex-document/src/index.ts";
import { fetchImage, readImageFile } from "./documentLibraryApi";

const MIME: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
  bmp: "image/bmp",
  svg: "image/svg+xml",
  avif: "image/avif",
};

/** A path or address as the text gives it, the way images.ts asks for it. */
export function imagePathKey(src: string): string {
  let value = src.trim().replace(/^<|>$/g, "");
  try {
    value = decodeURI(value);
  } catch {
    // as written
  }
  return value.replace(/^\.\//, "");
}

function sources(text: string, kind: "markdown" | "latex"): string[] {
  const found = new Set<string>();
  if (kind === "markdown") {
    for (const m of text.matchAll(/!\[[^\]\n]*\]\(\s*(<[^>\n]+>|[^)\s]+)(?:\s+(?:"[^"]*"|'[^']*'|\([^)]*\)))?\s*\)/g)) found.add(m[1]);
    for (const m of text.matchAll(/<img\b[^>]*?\bsrc\s*=\s*(?:"([^"]+)"|'([^']+)'|([^\s>]+))/gi)) found.add(m[1] ?? m[2] ?? m[3]);
    // reference-style pictures: [id]: path
    for (const m of text.matchAll(/^\s{0,3}\[[^\]\n]+\]:\s*(<[^>\n]+>|\S+\.(?:png|jpe?g|gif|webp|bmp|svg|avif))\b/gim)) found.add(m[1]);
  } else {
    for (const m of text.matchAll(/\\includegraphics\s*(?:\[[^\]]*\])?\s*\{([^}]+)\}/g)) found.add(m[1].trim());
  }
  return [...found].filter((src) => !/^data:/i.test(src));
}

const isAbsolute = (path: string) => /^(?:[a-zA-Z]:[\\/]|\/|\\\\)/.test(path);

/** path -> image key for the pictures `text` names that could be found. */
export async function loadImportImages(
  text: string,
  kind: "markdown" | "latex",
  where: { archive: Map<string, Uint8Array> | null; base: string; folder: string | null; folders: string[] },
): Promise<Map<string, string>> {
  const keys = new Map<string, string>();
  const extensions = kind === "latex" ? ["", ".png", ".jpg", ".jpeg", ".PNG", ".JPG"] : [""];
  await Promise.all(
    sources(text, kind).map(async (src) => {
      const path = imagePathKey(src);
      let bytes: Uint8Array | null = null;
      let name = path;
      if (/^https?:\/\//i.test(path)) bytes = await fetchImage(src.replace(/^<|>$/g, ""));
      else {
        const local = path.replace(/^file:\/\/(?=\/|[a-zA-Z]:)/i, "").replace(/^\/(?=[a-zA-Z]:)/, "");
        for (const folder of ["", ...where.folders]) {
          for (const extension of extensions) {
            if (bytes) break;
            const relative = folder + local + extension;
            if (where.archive) {
              const inArchive = (where.base + relative).replace(/\/\.\//g, "/").replace(/^\.\//, "");
              bytes = where.archive.get(inArchive) ?? where.archive.get(relative) ?? null;
            } else if (where.folder || isAbsolute(local)) {
              bytes = await readImageFile(isAbsolute(local) ? local + extension : where.folder + relative);
            }
            if (bytes) name = relative;
          }
        }
      }
      if (!bytes) return;
      const key =
        registerImageBytes(bytes) ??
        (await registerImageFile(new Blob([bytes as BlobPart], { type: MIME[/\.([a-z0-9]+)(?:[?#].*)?$/i.exec(name)?.[1]?.toLowerCase() ?? ""] ?? "" })));
      if (key) keys.set(path, key);
    }),
  );
  return keys;
}
