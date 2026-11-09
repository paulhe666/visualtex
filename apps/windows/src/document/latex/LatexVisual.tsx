// What the visual editor draws on the pages (visualEditor.ts): the caret, the selection, the text being composed
// by an input method, the formula editor in the room TeX left for it, the "\" menu; and the keyboard's sink — a
// textarea kept at the caret (an input method's candidates show there) that turns keys and text into edits.
import { useEffect, useLayoutEffect, useMemo, useRef, useState, type MutableRefObject } from "react";
import { DOCUMENT_COMMANDS, type CommandEntry, type FormulaSessionCallbacks } from "../../../../shared/visualtex-document/src/index.ts";
import { Preview } from "../CommandMenu";
import { DocumentFormulaField, type DocumentFormulaFieldApi } from "../DocumentFormulaField";
import { PENDING_FONT, type CaretView, type VisualEditor } from "./visualEditor";

/** The caret, the selection and the composition on page `page` (bp × zoom). */
export function PageMarks({ editor, page, zoom }: { editor: VisualEditor; page: number; zoom: number }) {
  const caret = editor.caretView();
  const rects = editor.selectionView().filter((r) => r.page === page);
  const pending = editor.pendingView(page);
  return (
    <>
      {pending.map((p, i) => (
        <span key={`p${i}`}>
          {p.covered.map((r, k) => (
            <div key={k} className="vtx-tex-pending-cover" style={{ left: r.x * zoom, top: r.y * zoom, width: r.w * zoom, height: r.h * zoom }} />
          ))}
          {p.text && (
            <span
              className="vtx-tex-pending"
              style={{ left: p.x * zoom, top: p.top * zoom, height: (p.bottom - p.top) * zoom, lineHeight: `${(p.bottom - p.top) * zoom}px`, fontSize: p.size * zoom, fontFamily: PENDING_FONT }}
            >
              {p.text}
            </span>
          )}
        </span>
      ))}
      {rects.map((r, i) => (
        <div key={i} className="vtx-tex-sel" style={{ left: r.x * zoom, top: r.y * zoom, width: r.w * zoom, height: r.h * zoom }} />
      ))}
      {caret && caret.page === page && <CaretMark caret={caret} zoom={zoom} composing={editor.composing} version={editor.version} />}
    </>
  );
}

function CaretMark({ caret, zoom, composing, version }: { caret: CaretView; zoom: number; composing: string; version: number }) {
  const height = (caret.bottom - caret.top) * zoom;
  return (
    <>
      {composing && (
        <span className="vtx-tex-compose" style={{ left: caret.x * zoom, top: caret.top * zoom, height, fontSize: height * 0.72, lineHeight: `${height}px` }}>
          {composing}
        </span>
      )}
      {/* (keyed by the editor's version: the blink starts again where the caret moved) */}
      <div key={version} className="vtx-tex-caret" style={{ left: caret.x * zoom - 0.75, top: caret.top * zoom, height }} />
    </>
  );
}

/**
 * The keyboard's textarea, at the caret on its page (in the pages' scroller): keys that move or edit go to the
 * editor; typed text (an input method's included, when it is committed) goes in as text.
 */
export function KeyboardSink({
  editor,
  zoom,
  scroller,
  sinkRef,
  onCopy,
}: {
  editor: VisualEditor;
  zoom: number;
  scroller: HTMLElement | null;
  sinkRef: MutableRefObject<HTMLTextAreaElement | null>;
  onCopy(text: { latex: string; plain: string }): void;
}) {
  const caret = editor.caretView();
  const [at, setAt] = useState<{ left: number; top: number; height: number } | null>(null);
  useLayoutEffect(() => {
    if (!caret || !scroller) return;
    const frame = scroller.querySelector<HTMLElement>(`.vtx-tex-page[data-page="${caret.page}"]`);
    if (!frame) return;
    const next = { left: frame.offsetLeft + caret.x * zoom, top: frame.offsetTop + caret.top * zoom, height: (caret.bottom - caret.top) * zoom };
    if (!at || Math.abs(at.left - next.left) > 0.5 || Math.abs(at.top - next.top) > 0.5 || Math.abs(at.height - next.height) > 0.5) setAt(next);
  });
  // a caret moved by the keyboard stays in view (the pages scroll to it)
  useLayoutEffect(() => {
    if (!at || !editor.keyMoved || !scroller) return;
    editor.keyMoved = false;
    const box = scroller.getBoundingClientRect();
    const top = at.top - scroller.scrollTop;
    const margin = 40;
    if (top < margin) scroller.scrollTop += top - margin;
    else if (top + at.height > box.height - margin) scroller.scrollTop += top + at.height - (box.height - margin);
  }, [at, editor, scroller]);
  const composing = useRef(false);
  const mac = useMemo(() => /Mac|iPhone|iPad/.test(navigator.platform), []);
  return (
    <textarea
      ref={sinkRef}
      className="vtx-tex-sink"
      aria-label="LaTeX"
      autoCapitalize="off"
      autoCorrect="off"
      spellCheck={false}
      style={at ? { left: at.left, top: at.top, height: at.height, fontSize: Math.max(12, at.height * 0.7) } : { left: 0, top: 0 }}
      onKeyDown={(e) => {
        if (composing.current || e.nativeEvent.isComposing) return;
        const mod = mac ? e.metaKey : e.ctrlKey;
        const menu = editor.menu;
        if (menu) {
          if (e.key === "Escape") {
            e.preventDefault();
            editor.closeMenu(true);
            return;
          }
          if (e.key === "Backspace") {
            e.preventDefault();
            editor.menuBack();
            return;
          }
          if (e.key === "ArrowDown" || e.key === "ArrowUp") {
            e.preventDefault();
            menu.index += e.key === "ArrowDown" ? 1 : -1;
            editor.menuType("");
            return;
          }
          if (e.key === "Enter" || e.key === "Tab") {
            e.preventDefault();
            const entries = commandEntries(menu.query);
            const entry = entries[Math.max(0, Math.min(menu.index, entries.length - 1))];
            if (entry) editor.runCommand(entry.id, entry.environment);
            else editor.closeMenu(true);
            return;
          }
          if (/^[1-9]$/.test(e.key) && !/\{/.test(menu.query)) {
            const entries = commandEntries(menu.query);
            const page = menu.index - (menu.index % 9);
            const entry = entries[page + Number(e.key) - 1];
            if (entry) {
              e.preventDefault();
              editor.runCommand(entry.id, entry.environment);
              return;
            }
          }
          if (e.key === " ") {
            e.preventDefault();
            editor.closeMenu(true);
            return;
          }
          return;
        }
        // (a click being placed: keys that move or delete wait with it — dropped, the caret is not there yet)
        if (editor.locating && /^(Arrow|Home|End|Backspace|Delete|Enter|Tab)/.test(e.key)) {
          e.preventDefault();
          return;
        }
        const handled = (() => {
          switch (e.key) {
            case "ArrowLeft":
              editor.move(-1, e.shiftKey);
              return true;
            case "ArrowRight":
              editor.move(1, e.shiftKey);
              return true;
            case "ArrowUp":
              editor.moveLine(-1, e.shiftKey);
              return true;
            case "ArrowDown":
              editor.moveLine(1, e.shiftKey);
              return true;
            case "Home":
              editor.lineEdge(-1, e.shiftKey);
              return true;
            case "End":
              editor.lineEdge(1, e.shiftKey);
              return true;
            case "Backspace":
              editor.deleteChar(-1);
              return true;
            case "Delete":
              editor.deleteChar(1);
              return true;
            case "Enter":
              if (mod) return false;
              editor.newParagraph(e.shiftKey);
              return true;
            case "Tab":
              return true;
            case "Escape":
              if (editor.caret) editor.caret.anchor = editor.caret.offset;
              editor.menuType("");
              return true;
          }
          if (!mod) return false;
          switch (e.key.toLowerCase()) {
            case "z":
              editor.history(e.shiftKey ? "redo" : "undo");
              return true;
            case "y":
              if (mac) return false;
              editor.history("redo");
              return true;
            case "a":
              editor.selectAll();
              return true;
            case "b":
              editor.wrap("textbf");
              return true;
            case "i":
              editor.wrap("textit");
              return true;
            case "u":
              editor.wrap("underline");
              return true;
            case "e":
              editor.newFormula(false);
              return true;
          }
          return false;
        })();
        if (handled) e.preventDefault();
      }}
      onCompositionStart={() => {
        composing.current = true;
        editor.setComposing("");
      }}
      onCompositionUpdate={(e) => {
        editor.setComposing(e.data);
      }}
      onCompositionEnd={(e) => {
        composing.current = false;
        editor.setComposing("");
        const text = e.data;
        e.currentTarget.value = "";
        if (text) {
          if (editor.menu) editor.menuType(text);
          else editor.insertText(text);
        } else editor.menuType("");
      }}
      onInput={(e) => {
        if (composing.current || (e.nativeEvent as InputEvent).isComposing) return;
        const text = e.currentTarget.value;
        e.currentTarget.value = "";
        if (!text) return;
        if (editor.menu) {
          if (/^[A-Za-z@*{}]+$/.test(text)) editor.menuType(text);
          else {
            editor.closeMenu(true);
            editor.insertText(text);
          }
          return;
        }
        editor.insertText(text);
      }}
      onCopy={(e) => {
        const out = editor.copyText();
        if (!out) return;
        e.preventDefault();
        e.clipboardData.setData("text/plain", out.latex);
        onCopy(out);
      }}
      onCut={(e) => {
        const out = editor.cut();
        if (!out) return;
        e.preventDefault();
        e.clipboardData.setData("text/plain", out.latex);
      }}
      onPaste={(e) => {
        const text = e.clipboardData.getData("text/plain");
        e.preventDefault();
        if (text) editor.paste(text);
      }}
    />
  );
}

/** The "\" menu's entries for what is typed after the backslash (the block editor's LaTeX commands). */
export function commandEntries(query: string): CommandEntry[] {
  const q = query.toLowerCase();
  if (!q) return DOCUMENT_COMMANDS.filter((e) => e.featured);
  const begin = /^begin\{?([^}]*)/.exec(q);
  const name = begin ? begin[1] : q;
  return DOCUMENT_COMMANDS.filter((e) =>
    begin ? !!e.environment && e.environment.toLowerCase().startsWith(name) : e.names.some((n) => n.toLowerCase().startsWith(name)) || e.label.toLowerCase().includes("\\" + name),
  ).slice(0, 40);
}

/** The "\" menu under the caret on its page. */
export function LatexCommandMenu({ editor, page, zoom, isEn }: { editor: VisualEditor; page: number; zoom: number; isEn: boolean }) {
  const menu = editor.menu;
  const caret = editor.caretView();
  const entries = useMemo(() => (menu ? commandEntries(menu.query) : []), [menu?.query]);
  const listRef = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    listRef.current?.querySelector(".is-current")?.scrollIntoView({ block: "nearest" });
  });
  if (!menu || !caret || caret.page !== page) return null;
  const index = entries.length ? ((menu.index % entries.length) + entries.length) % entries.length : -1;
  const pageStart = index - (index % 9);
  return (
    <div
      className="vtx-cmd-menu vtx-tex-cmd-menu"
      style={{ left: caret.x * zoom - 10, top: caret.bottom * zoom + 6 }}
      role="listbox"
      aria-label={isEn ? "LaTeX commands" : "LaTeX 命令"}
      onPointerDown={(e) => {
        e.preventDefault();
        e.stopPropagation();
      }}
    >
      <div className="vtx-tex-cmd-query">\{menu.query}</div>
      <div className="vtx-cmd-list" ref={listRef}>
        {entries.map((entry, i) => (
          <div
            key={entry.id}
            role="option"
            aria-selected={i === index}
            className={"vtx-cmd-item" + (i === index ? " is-current" : "")}
            onMouseEnter={() => {
              if (menu.index !== i) {
                menu.index = i;
                editor.menuType("");
              }
            }}
            onClick={() => editor.runCommand(entry.id, entry.environment)}
          >
            <span className="vtx-cmd-index">{i >= pageStart && i < pageStart + 9 ? i - pageStart + 1 : ""}</span>
            <span className="vtx-cmd-copy">
              <span className="vtx-cmd-label">{entry.label}</span>
              <span className="vtx-cmd-desc">{isEn ? entry.description.en : entry.description.zh}</span>
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

/**
 * The formula editor on page `page`, in the room TeX left for the formula (its baseline on the line's): its box
 * is measured and given to TeX, so the text around it makes room as for the formula.
 */
export function FormulaOnPage({ editor, page, zoom, onClosed }: { editor: VisualEditor; page: number; zoom: number; onClosed(): void }) {
  const f = editor.formula;
  const wrapRef = useRef<HTMLDivElement>(null);
  const probeRef = useRef<HTMLSpanElement>(null);
  const apiRef = useRef<DocumentFormulaFieldApi>(null);
  const [baseline, setBaseline] = useState(0);
  /** a click's point in the formula: its caret placed there once the editor stands where the formula is */
  const placed = useRef<string | null>(null);
  const callbacks = useMemo<FormulaSessionCallbacks>(
    () => ({
      edit: (latex) => editor.formulaEdited(latex),
      exit: (how) => {
        editor.closeFormula(how === "before" || how === "up" ? "before" : "after");
        onClosed();
      },
      remove: () => {
        editor.closeFormula("remove");
        onClosed();
      },
      dollar: () => {
        editor.closeFormula("after");
        onClosed();
      },
      undo: () => false,
      redo: () => false,
      handBack: (keys) => {
        for (const k of keys) if (k.text) editor.insertText(k.text);
      },
    }),
    [editor, onClosed],
  );
  // the editor's box: its width, height and depth around the baseline (the probe sits on it), to TeX in bp
  useLayoutEffect(() => {
    const wrap = wrapRef.current;
    const probe = probeRef.current;
    if (!wrap || !probe) return;
    const measure = () => {
      const field = wrap.querySelector<HTMLElement>(".vtx-tex-formula-field");
      if (!field) return;
      const box = field.getBoundingClientRect();
      const base = probe.getBoundingClientRect().top;
      const offset = base - wrap.getBoundingClientRect().top;
      setBaseline(offset);
      if (box.width > 0) editor.formulaSized(box.width / zoom, Math.max(0, base - box.top) / zoom, Math.max(0, box.bottom - base) / zoom);
      // the editor drawn, on the line's baseline: the caret where the click was
      const key = f ? `${f.unit.path}:${f.at}` : null;
      if (f && f.entry.kind === "point" && box.height > 2 && Math.abs(offset - baseline) < 0.5 && placed.current !== key) {
        placed.current = key;
        const entry = f.entry;
        requestAnimationFrame(() => apiRef.current?.enter(entry));
      }
    };
    measure();
    const ro = new ResizeObserver(() => requestAnimationFrame(measure));
    const field = wrap.querySelector<HTMLElement>(".vtx-tex-formula-field");
    if (field) ro.observe(field);
    return () => ro.disconnect();
  });
  useEffect(() => {
    // focused once it is there
    const t = window.setTimeout(() => apiRef.current?.focus(), 0);
    return () => window.clearTimeout(t);
  }, [f?.unit, f?.at]);
  if (!f || !f.place || f.place.page !== page) return null;
  return (
    <div
      ref={wrapRef}
      className={"vtx-tex-formula-on-page" + (f.display ? " is-display" : "")}
      style={{ left: f.place.x * zoom, top: f.place.y * zoom - baseline }}
      onPointerDown={(e) => e.stopPropagation()}
      onMouseDown={(e) => e.stopPropagation()}
    >
      <span ref={probeRef} className="vtx-tex-formula-probe" />
      <span className="vtx-tex-formula-field vtx-formula-editor">
        <DocumentFormulaField
          id={`latex-page-formula-${f.unit.path}-${f.at}`}
          latex={f.latex}
          display={f.display}
          entry={f.entry}
          fontSize={f.fontSize * zoom}
          callbacks={callbacks}
          apiRef={apiRef}
        />
      </span>
    </div>
  );
}
