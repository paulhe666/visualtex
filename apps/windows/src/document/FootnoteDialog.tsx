// The footnote editor (DESIGN.md §8.3): opened by a new footnote, a click on a footnote's number or on its note at
// the foot of the page. The note's text, $…$ for formulas; Enter applies (Shift+Enter: a new line), Esc closes,
// "Delete footnote" removes the mark with its note.
import { useLayoutEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import type { EditorView } from "prosemirror-view";
import { setFootnoteText } from "../../../shared/visualtex-document/src/index.ts";

type Lang = "en" | "zh";
const t = (lang: Lang, en: string, zh: string) => (lang === "en" ? en : zh);

export function FootnoteDialog(props: { view: EditorView; pos: number; rect: DOMRect | null; lang: Lang; onClose: () => void }) {
  const { view, pos, lang, onClose } = props;
  // the footnote is the node at `pos` when the dialog opens; it is found again by its id when the text is applied
  const [target] = useState(() => {
    const node = view.state.doc.nodeAt(pos);
    return node?.type.name === "footnote" ? { id: node.attrs.id as string | null, text: node.attrs.text as string } : null;
  });
  const [text, setText] = useState(target?.text ?? "");
  const panelRef = useRef<HTMLFormElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const [place, setPlace] = useState<{ left: number; top: number } | null>(null);

  useLayoutEffect(() => {
    const panel = panelRef.current;
    if (!panel) return;
    const anchor = props.rect ?? (() => {
      try {
        const c = view.coordsAtPos(pos);
        return { left: c.left, top: c.top, bottom: c.bottom } as DOMRect;
      } catch {
        return { left: 40, top: 80, bottom: 80 } as DOMRect;
      }
    })();
    // under the anchor, else above it; always whole inside the window
    const height = panel.offsetHeight;
    const below = anchor.bottom + 6;
    const preferred = below + height <= window.innerHeight - 4 ? below : anchor.top - height - 6;
    const top = Math.max(4, Math.min(preferred, window.innerHeight - height - 4));
    setPlace({ left: Math.max(4, Math.min(anchor.left - 12, window.innerWidth - panel.offsetWidth - 4)), top });
    inputRef.current?.focus({ preventScroll: true });
    inputRef.current?.select();
  }, [view, pos, props.rect]);

  /** The footnote's position now (the document may have changed around it). */
  const currentPos = (): number | null => {
    if (!target) return null;
    const at = view.state.doc.nodeAt(pos);
    if (at?.type.name === "footnote" && at.attrs.id === target.id) return pos;
    let found: number | null = null;
    view.state.doc.descendants((node, p) => {
      if (found !== null) return false;
      if (node.type.name === "footnote" && node.attrs.id === target.id) found = p;
      return true;
    });
    return found;
  };
  const close = () => {
    onClose();
    view.focus();
  };
  const apply = () => {
    const at = currentPos();
    if (at !== null) setFootnoteText(at, text.trim())(view.state, view.dispatch);
    close();
  };
  const remove = () => {
    const at = currentPos();
    if (at !== null) view.dispatch(view.state.tr.delete(at, at + 1).scrollIntoView());
    close();
  };

  if (!target) return null;
  return createPortal(
    <>
      <div className="document-control-dismiss" onMouseDown={apply} />
      <form
        ref={panelRef}
        className="document-control-panel vtx-link-dialog vtx-footnote-dialog"
        role="dialog"
        aria-label={t(lang, "Footnote", "脚注")}
        style={place ? { left: place.left, top: place.top } : { left: -9999, top: -9999 }}
        onSubmit={(event) => {
          event.preventDefault();
          apply();
        }}
        onKeyDown={(event) => {
          if (event.key === "Escape") {
            event.preventDefault();
            close();
          } else if (event.key === "Enter" && !event.shiftKey && !event.nativeEvent.isComposing) {
            event.preventDefault();
            apply();
          }
        }}
      >
        <div className="vtx-footnote-dialog-title">{t(lang, "Footnote", "脚注")}</div>
        <textarea
          ref={inputRef}
          value={text}
          rows={3}
          onChange={(event) => setText(event.target.value)}
          spellCheck={false}
        />
        <div className="vtx-link-dialog-actions">
          <button type="button" className="secondary-button" onClick={remove}>
            {t(lang, "Delete footnote", "删除脚注")}
          </button>
          <span className="vtx-link-dialog-hint" />
          <button type="submit" className="primary-small-button">
            {t(lang, "OK", "确定")}
          </button>
        </div>
      </form>
    </>,
    document.body,
  );
}
