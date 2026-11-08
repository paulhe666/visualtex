// The menu under a LaTeX command typed in the document text (suggest.ts): VisualTeX's own command list, drawn as
// the formula editor's "\" suggestions are (numbered rows, the command, a picture of what it makes). The keys
// stay in the text: ↑/↓, Enter or Tab, 1-9, Esc are the editor's; a click runs an entry too.
import { useLayoutEffect, useMemo, useRef, useState } from "react";
import { Image as ImageIcon } from "lucide-react";
import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import {
  commandMenuEntries,
  commandToken,
  highlightCommandEntry,
  runCommandEntry,
  type CommandEntry,
  type CommandPreview,
} from "../../shared/visualtex-document/src/index.ts";
import { cachedPreviewMarkup } from "../components/MathPreview";
import { markVisualTexFormulaFontGlyphs } from "../editor/formulaFontPreferences";
import { getFormulaFontRevision } from "../editor/visualTexLayoutRuntime";
import { getCustomSymbolRevision } from "../math/customSymbolRegistry";

type Lang = "en" | "zh";

const PAGE = 9;

function formulaMarkup(latex: string): string {
  const markup = cachedPreviewMarkup(latex, getCustomSymbolRevision(), getFormulaFontRevision());
  if (!markup) return "";
  const holder = document.createElement("span");
  holder.innerHTML = markup;
  markVisualTexFormulaFontGlyphs(holder);
  return holder.innerHTML;
}

function Formula({ latex, display = false }: { latex: string; display?: boolean }) {
  const html = useMemo(() => formulaMarkup(display ? latex : `{\\textstyle ${latex}}`), [latex, display]);
  return <span className="vtx-cmd-formula" dangerouslySetInnerHTML={{ __html: html }} />;
}

function Preview({ preview }: { preview: CommandPreview }) {
  switch (preview.kind) {
    case "math":
      return <Formula latex={preview.latex} />;
    case "display":
      return (
        <span className="vtx-cmd-display">
          <Formula latex={preview.latex} display />
          {preview.numbered && <span className="vtx-cmd-tag">(1)</span>}
        </span>
      );
    case "text":
      return (
        <span
          className={"vtx-cmd-sample" + (preview.style ? ` is-${preview.style}` : "")}
          style={preview.scale ? { fontSize: `${Math.min(26, Math.max(7, 13 * preview.scale))}px` } : undefined}
        >
          {preview.text}
        </span>
      );
    case "section":
      return (
        <span className={`vtx-cmd-section is-level-${preview.level}`}>
          {preview.unnumbered ? null : ["1", "1.1", "1.1.1"][preview.level - 1]} <span className="vtx-cmd-line" />
        </span>
      );
    case "title":
      return (
        <span className="vtx-cmd-title">
          <span className="vtx-cmd-line" />
        </span>
      );
    case "list":
      return (
        <span className="vtx-cmd-list-mock">
          {[1, 2].map((n) => (
            <span key={n}>
              <span className="vtx-cmd-marker">{preview.ordered ? `${n}.` : "•"}</span>
              <span className="vtx-cmd-line" />
            </span>
          ))}
        </span>
      );
    case "align":
      return (
        <span className={`vtx-cmd-align is-${preview.align}`}>
          <span className="vtx-cmd-line is-long" />
          <span className="vtx-cmd-line is-short" />
          <span className="vtx-cmd-line is-mid" />
        </span>
      );
    case "block":
      if (preview.block === "table") {
        return (
          <span className="vtx-cmd-table">
            {Array.from({ length: 6 }, (_, i) => (
              <span key={i} />
            ))}
          </span>
        );
      }
      if (preview.block === "rule") return <span className="vtx-cmd-rule" />;
      if (preview.block === "code") return <span className="vtx-cmd-code">{"{ }"}</span>;
      if (preview.block === "image") return <ImageIcon size={18} strokeWidth={1.6} aria-hidden="true" />;
      if (preview.block === "quote") {
        return (
          <span className="vtx-cmd-quote">
            <span className="vtx-cmd-line is-long" />
            <span className="vtx-cmd-line is-mid" />
          </span>
        );
      }
      return <span className="vtx-cmd-sample">↵</span>;
    case "ref":
      return <span className="vtx-cmd-sample">({preview.number ?? "n"})</span>;
  }
}

export function CommandMenu(props: { view: EditorView | null; state: EditorState | null; lang: Lang; frame: HTMLElement | null }) {
  const { view, state, lang, frame } = props;
  const token = state ? commandToken(state) : null;
  const entries = useMemo<CommandEntry[]>(() => (state && token ? commandMenuEntries(state) : []), [state]);
  const listRef = useRef<HTMLDivElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const [above, setAbove] = useState(false);
  // the menu's width (the candidate bar is as wide as its entries)
  const [width, setWidth] = useState(392);
  const index = token && entries.length ? Math.min(token.index, entries.length - 1) : -1;
  const visible = !!(view && frame && token && !token.dismissed && entries.length);

  useLayoutEffect(() => {
    listRef.current?.querySelector(".is-current")?.scrollIntoView({ block: "nearest" });
  }, [index, visible]);

  // Under the command, or above it when the window has no room below.
  const anchor = visible ? view!.coordsAtPos(Math.min(token!.from, view!.state.doc.content.size)) : null;
  useLayoutEffect(() => {
    if (!visible || !anchor || !menuRef.current) return;
    const height = menuRef.current.offsetHeight;
    const room = window.innerHeight - anchor.bottom - 8;
    setAbove(room < height && anchor.top - 8 > room);
    const measured = menuRef.current.offsetWidth + 12;
    if (Math.abs(measured - width) > 1) setWidth(measured);
  });
  if (!visible || !anchor) return null;

  const box = frame!.getBoundingClientRect();
  const left = Math.max(4, Math.min(anchor.left - box.left + frame!.scrollLeft - 10, frame!.scrollWidth - width));
  const top = above
    ? anchor.top - box.top + frame!.scrollTop - 6 - (menuRef.current?.offsetHeight ?? 0)
    : anchor.bottom - box.top + frame!.scrollTop + 6;
  const pageStart = index - (index % PAGE);
  // The numbered page (1-9 pick), as in the formula editor's suggestions; an argument being typed takes digits.
  const numbered = !/\{[^{}]*$/.test(token!.query) || /^begin\{/.test(token!.query);

  return (
    <div
      ref={menuRef}
      className="vtx-cmd-menu"
      style={{ left, top }}
      role="listbox"
      aria-label={token!.mode === "markdown" ? (lang === "en" ? "Markdown commands" : "Markdown 命令") : lang === "en" ? "LaTeX commands" : "LaTeX 命令"}
      onMouseDown={(event) => event.preventDefault()}
    >
      <div className="vtx-cmd-list" ref={listRef}>
        {entries.map((entry, i) => (
          <div
            key={entry.id}
            role="option"
            aria-selected={i === index}
            className={"vtx-cmd-item" + (i === index ? " is-current" : "")}
            onMouseEnter={() => i !== index && highlightCommandEntry(view!, i)}
            onClick={() => runCommandEntry(view!, entry, { via: "menu" })}
          >
            <span className="vtx-cmd-index">{numbered && i >= pageStart && i < pageStart + PAGE ? i - pageStart + 1 : ""}</span>
            <span className="vtx-cmd-copy">
              <span className="vtx-cmd-label">{entry.label}</span>
              <span className="vtx-cmd-desc">{lang === "en" ? entry.description.en : entry.description.zh}</span>
            </span>
            <span className="vtx-cmd-preview">
              <Preview preview={entry.preview} />
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}
