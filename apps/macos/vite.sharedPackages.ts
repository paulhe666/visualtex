// Packages that code in apps/shared imports: every build resolves them from this
// app's node_modules (vite's dedupe), since apps/shared has none of its own.
export const sharedPackageDedupe = [
  "react",
  "react-dom",
  "lucide-react",
  // the document editor (apps/shared/visualtex-document)
  "prosemirror-commands",
  "prosemirror-dropcursor",
  "prosemirror-gapcursor",
  "prosemirror-history",
  "prosemirror-inputrules",
  "prosemirror-keymap",
  "prosemirror-model",
  "prosemirror-schema-list",
  "prosemirror-state",
  "prosemirror-tables",
  "prosemirror-transform",
  "prosemirror-view",
  "prosemirror-markdown",
  "markdown-it",
  // the source editor's Markdown parser (sourceMarkdown.ts)
  "@lezer/markdown",
  "@lezer/common",
  "@lezer/highlight",
];
