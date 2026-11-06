// The link dialog (Ctrl+K, the menus' link entry): under the selection, the address and, for a new link without
// selected text, the text to show. On a link it changes or removes that link. Enter applies, Esc closes.
import { useLayoutEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import type { EditorView } from "prosemirror-view";
import { insertLink, linkAt, removeLink, setLink } from "../../shared/visualtex-document/src/index.ts";

type Lang = "en" | "zh";
const t = (lang: Lang, en: string, zh: string) => (lang === "en" ? en : zh);

export function LinkDialog(props: { view: EditorView; lang: Lang; onClose: () => void }) {
  const { view, lang, onClose } = props;
  // What the dialog works on is fixed when it opens.
  const [target] = useState(() => {
    const state = view.state;
    const existing = linkAt(state);
    const { from, to, empty } = state.selection;
    const selected = empty ? "" : state.doc.textBetween(from, to, " ");
    return { existing, from, to, empty, selected };
  });
  const urlLike = /^(https?:\/\/|www\.|mailto:)\S+$/i.test(target.selected.trim());
  const [href, setHref] = useState(target.existing?.href ?? (urlLike ? target.selected.trim() : ""));
  const [text, setText] = useState("");
  const needsText = !target.existing && target.empty;
  const panelRef = useRef<HTMLFormElement>(null);
  const hrefRef = useRef<HTMLInputElement>(null);
  const textRef = useRef<HTMLInputElement>(null);
  const [place, setPlace] = useState<{ left: number; top: number } | null>(null);

  useLayoutEffect(() => {
    const anchor = view.coordsAtPos(target.existing?.from ?? target.from);
    const panel = panelRef.current;
    if (!panel) return;
    const below = anchor.bottom + 6;
    const top = below + panel.offsetHeight <= window.innerHeight - 4 ? below : Math.max(4, anchor.top - panel.offsetHeight - 6);
    setPlace({ left: Math.max(4, Math.min(anchor.left, window.innerWidth - panel.offsetWidth - 4)), top });
    (needsText ? textRef : hrefRef).current?.focus();
  }, [view, target, needsText]);

  const close = () => {
    onClose();
    view.focus();
  };
  const apply = () => {
    const address = href.trim();
    if (!address) {
      if (target.existing) removeLink(view, target.existing.from, target.existing.to);
      close();
      return;
    }
    if (target.existing) setLink(view, target.existing.from, target.existing.to, address, target.existing.title);
    else if (!target.empty) setLink(view, target.from, target.to, address);
    else insertLink(view, text.trim(), address);
    close();
  };

  return createPortal(
    <>
      <div className="document-control-dismiss" onMouseDown={close} />
      <form
        ref={panelRef}
        className="document-control-panel vtx-link-dialog"
        role="dialog"
        aria-label={t(lang, "Link", "链接")}
        style={place ? { left: place.left, top: place.top } : { left: -9999, top: -9999 }}
        onSubmit={(event) => {
          event.preventDefault();
          apply();
        }}
        onKeyDown={(event) => {
          if (event.key === "Escape") {
            event.preventDefault();
            close();
          }
        }}
      >
        {needsText && (
          <label>
            <span>{t(lang, "Text", "文字")}</span>
            <input ref={textRef} value={text} onChange={(event) => setText(event.target.value)} />
          </label>
        )}
        <label>
          <span>{t(lang, "Address", "地址")}</span>
          <input
            ref={hrefRef}
            value={href}
            onChange={(event) => setHref(event.target.value)}
            placeholder="https://"
            spellCheck={false}
          />
        </label>
        <div className="vtx-link-dialog-actions">
          {target.existing && (
            <button
              type="button"
              className="secondary-button"
              onClick={() => {
                removeLink(view, target.existing!.from, target.existing!.to);
                close();
              }}
            >
              {t(lang, "Remove link", "移除链接")}
            </button>
          )}
          <span className="vtx-link-dialog-hint" />
          <button type="submit" className="primary-small-button">
            {target.existing ? t(lang, "Update", "更新") : t(lang, "Add link", "添加链接")}
          </button>
        </div>
      </form>
    </>,
    document.body,
  );
}
