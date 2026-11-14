import { fileURLToPath } from "node:url";

// The URL under which an app's Vite dev server serves a file of apps/shared/visualtex-app/src: Vite serves files
// outside the app folder as /@fs/<absolute path> (both apps allow ../shared in server.fs.allow). Browser
// regressions that import such a module inside the page use this instead of /src/<path>.
export function sharedSourceUrl(path) {
  const file = fileURLToPath(new URL(`../src/${path}`, import.meta.url)).replace(/\\/g, "/");
  return `/@fs${file.startsWith("/") ? "" : "/"}${file}`;
}
