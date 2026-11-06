// Code block highlighting as VS Code does it: Shiki runs VS Code's own TextMate grammars with its Dark+ theme
// (docs/document-editor/DESIGN.md §6.5). Shiki and each grammar load on first use, each in its own chunk; until a
// grammar is in, its blocks show plain text, then the document draws them again (onReady).
import type { HighlighterCore, LanguageRegistration } from "shiki/core";
import type { CodeHighlighter, CodeToken } from "../../shared/visualtex-document/src/index.ts";

type Grammar = () => Promise<{ default: LanguageRegistration[] }>;

/** The block's languages -> Shiki's grammar (and Shiki's name for it). */
const GRAMMARS: Record<string, { name: string; load: Grammar }> = {
  python: { name: "python", load: () => import("shiki/langs/python.mjs") },
  javascript: { name: "javascript", load: () => import("shiki/langs/javascript.mjs") },
  typescript: { name: "typescript", load: () => import("shiki/langs/typescript.mjs") },
  c: { name: "c", load: () => import("shiki/langs/c.mjs") },
  cpp: { name: "cpp", load: () => import("shiki/langs/cpp.mjs") },
  csharp: { name: "csharp", load: () => import("shiki/langs/csharp.mjs") },
  java: { name: "java", load: () => import("shiki/langs/java.mjs") },
  kotlin: { name: "kotlin", load: () => import("shiki/langs/kotlin.mjs") },
  swift: { name: "swift", load: () => import("shiki/langs/swift.mjs") },
  go: { name: "go", load: () => import("shiki/langs/go.mjs") },
  rust: { name: "rust", load: () => import("shiki/langs/rust.mjs") },
  matlab: { name: "matlab", load: () => import("shiki/langs/matlab.mjs") },
  julia: { name: "julia", load: () => import("shiki/langs/julia.mjs") },
  r: { name: "r", load: () => import("shiki/langs/r.mjs") },
  fortran: { name: "fortran-free-form", load: () => import("shiki/langs/fortran-free-form.mjs") },
  haskell: { name: "haskell", load: () => import("shiki/langs/haskell.mjs") },
  scala: { name: "scala", load: () => import("shiki/langs/scala.mjs") },
  ruby: { name: "ruby", load: () => import("shiki/langs/ruby.mjs") },
  php: { name: "php", load: () => import("shiki/langs/php.mjs") },
  lua: { name: "lua", load: () => import("shiki/langs/lua.mjs") },
  bash: { name: "shellscript", load: () => import("shiki/langs/shellscript.mjs") },
  powershell: { name: "powershell", load: () => import("shiki/langs/powershell.mjs") },
  sql: { name: "sql", load: () => import("shiki/langs/sql.mjs") },
  json: { name: "json", load: () => import("shiki/langs/json.mjs") },
  yaml: { name: "yaml", load: () => import("shiki/langs/yaml.mjs") },
  toml: { name: "toml", load: () => import("shiki/langs/toml.mjs") },
  xml: { name: "xml", load: () => import("shiki/langs/xml.mjs") },
  html: { name: "html", load: () => import("shiki/langs/html.mjs") },
  css: { name: "css", load: () => import("shiki/langs/css.mjs") },
  latex: { name: "latex", load: () => import("shiki/langs/latex.mjs") },
  markdown: { name: "markdown", load: () => import("shiki/langs/markdown.mjs") },
  dockerfile: { name: "dockerfile", load: () => import("shiki/langs/dockerfile.mjs") },
  ini: { name: "ini", load: () => import("shiki/langs/ini.mjs") },
};

const THEME = "dark-plus";
const CACHE_LIMIT = 300;

let core: Promise<HighlighterCore> | null = null;
let ready: HighlighterCore | null = null;
const loaded = new Set<string>();
const loading = new Set<string>();
const listeners = new Set<() => void>();
const cache = new Map<string, CodeToken[][]>();

function shiki(): Promise<HighlighterCore> {
  core ??= (async () => {
    const [{ createHighlighterCore }, { createJavaScriptRegexEngine }, theme] = await Promise.all([
      import("shiki/core"),
      import("shiki/engine/javascript"),
      import("shiki/themes/dark-plus.mjs"),
    ]);
    return createHighlighterCore({ themes: [theme.default], langs: [], engine: createJavaScriptRegexEngine({ forgiving: true }) });
  })();
  return core;
}

function load(language: string) {
  const grammar = GRAMMARS[language];
  if (!grammar || loaded.has(language) || loading.has(language)) return;
  loading.add(language);
  void (async () => {
    try {
      const highlighter = await shiki();
      const module = await grammar.load();
      await highlighter.loadLanguage(...module.default);
      ready = highlighter;
      loaded.add(language);
      for (const listener of listeners) listener();
    } catch {
      // The language stays plain text.
    } finally {
      loading.delete(language);
    }
  })();
}

export const codeHighlighter: CodeHighlighter = {
  tokens(code, language) {
    const grammar = GRAMMARS[language];
    if (!grammar) return null;
    if (!ready || !loaded.has(language)) {
      load(language);
      return null;
    }
    const key = `${language}\u0000${code}`;
    const cached = cache.get(key);
    if (cached) return cached;
    const lines = ready.codeToTokensBase(code, { lang: grammar.name, theme: THEME });
    const tokens = lines.map((line) => line.map((token) => ({ content: token.content, color: token.color, fontStyle: token.fontStyle })));
    cache.set(key, tokens);
    if (cache.size > CACHE_LIMIT) cache.delete(cache.keys().next().value as string);
    return tokens;
  },
  onReady(listener) {
    listeners.add(listener);
    return () => {
      listeners.delete(listener);
    };
  },
};
