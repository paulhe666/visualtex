// LaTeX documents' TeX side (personal edition): one interface, two ways to it. In the app the tex_* Tauri
// commands (the tex-bridge crate in src-tauri); in a browser on the dev server, the same commands through
// vite.texBridge.ts (POST /__vtx_tex). Elsewhere there is no TeX.
import { invoke, isTauri } from "@tauri-apps/api/core";

export interface TexDist {
  bin: string;
  version: string;
  engines: string[];
  dvisvgm: boolean;
  biber: boolean;
  mutool: string | null;
  gs: string | null;
}

export interface TexStatus {
  found: boolean;
  dist: TexDist | null;
}

export interface ProjectFile {
  path: string;
  size: number;
  /** "tex" | "bib" | "image" | "other" */
  kind: string;
}

export interface ProjectInfo {
  root: string;
  main: string | null;
  files: ProjectFile[];
  engine: string | null;
}

export interface PageMeta {
  width: number;
  height: number;
  hash: string;
  unsupported: string[];
}

export interface SvgPage {
  page: number;
  hash: string;
  svg: string;
}

export interface CompileTimings {
  sync: number;
  tex: number;
  bib: number;
  screen: number;
  svg: number;
  total: number;
  passes: number;
}

export interface CompileResult {
  status: "ok" | "failed" | "superseded" | "notex";
  build: number;
  engine: string;
  engineReason: string;
  pageMode: "svg" | "raster";
  pages: PageMeta[];
  svgs: SvgPage[];
  log: string;
  needsRerun: boolean;
  rerunReason: string;
  warm: boolean;
  timings: CompileTimings;
  error: string;
}

export interface CompileRequest {
  doc: string;
  root: string;
  main: string;
  engine?: string | null;
  shellEscape?: boolean;
  overlays?: Record<string, string>;
  visible?: number[];
  full?: boolean;
  /** names the compile for preview() */
  ticket?: string;
}

export interface PreviewResult {
  /** "early": the visible pages of a run still going; "done": nothing before the result */
  status: "early" | "done";
  pages: PageMeta[];
  svgs: SvgPage[];
  at: number;
}

export interface PageOut {
  page: number;
  hash: string;
  svg: string | null;
  png: string | null;
}

export interface PdfOut {
  path: string;
  size: number;
  base64: string | null;
}

export interface SynctexOut {
  base64: string;
  work: string;
  aliases: string[];
}

export interface TexService {
  status(): Promise<TexStatus>;
  /** a folder, or a .tex file (its folder; it the main file) */
  open(root: string, main?: string | null): Promise<ProjectInfo>;
  /** an empty project folder made by the app */
  create(name: string): Promise<ProjectInfo>;
  files(root: string): Promise<ProjectInfo>;
  /** a file's bytes (base64) */
  read(root: string, path: string): Promise<string>;
  write(root: string, path: string, content: { text: string } | { base64: string }): Promise<void>;
  compile(request: CompileRequest): Promise<CompileResult>;
  /** the visible pages of compile `ticket` as soon as TeX has written them */
  preview(doc: string, ticket: string): Promise<PreviewResult>;
  pages(doc: string, build: number, pages: number[], scale?: number): Promise<PageOut[]>;
  pdf(doc: string, build: number, options: { dest?: string; bytes?: boolean }): Promise<PdfOut>;
  synctex(doc: string, build: number): Promise<SynctexOut>;
  close(doc: string): Promise<void>;
}

function make(call: <T>(command: string, request: unknown) => Promise<T>): TexService {
  return {
    status: () => call("tex_status", {}),
    open: (root, main) => call("tex_project_open", { root, main: main ?? null }),
    create: (name) => call("tex_project_new", { name }),
    files: (root) => call("tex_project_files", { root }),
    read: (root, path) => call("tex_project_read", { root, path }),
    write: (root, path, content) => call("tex_project_write", { root, path, ...content }),
    compile: (request) => call("tex_compile", request),
    preview: (doc, ticket) => call("tex_preview", { doc, ticket }),
    pages: (doc, build, pages, scale) => call("tex_pages", { doc, build, pages, scale }),
    pdf: (doc, build, options) => call("tex_pdf", { doc, build, ...options }),
    synctex: (doc, build) => call("tex_synctex", { doc, build }),
    close: (doc) => call("tex_close", { doc }),
  };
}

const tauriTex = make((command, request) => invoke(command, { request }));

const devTex = make(async <T,>(command: string, request: unknown): Promise<T> => {
  const response = await fetch("/__vtx_tex", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ cmd: command, args: { request } }),
  });
  if (!response.ok) throw new Error(`TeX bridge: HTTP ${response.status}`);
  const reply = (await response.json()) as { ok: boolean; result?: T; error?: string };
  if (!reply.ok) throw new Error(reply.error ?? "TeX bridge error");
  return reply.result as T;
});

/** The TeX side of this build (null: none — a public or production web build). */
export function texService(): TexService | null {
  if (isTauri()) return tauriTex;
  if (import.meta.env.DEV) return devTex;
  return null;
}

export function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}
