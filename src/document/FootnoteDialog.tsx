// The footnote editor (DESIGN.md §8.3): opened by a new footnote, a click on a footnote's number or on its note at
// the foot of the page. The note is edited as the document is: text and inline formulas ("$$", "\(", \frac… open a
// formula in place); Enter applies (Shift+Enter: a new line), Esc closes, "Delete footnote" removes the mark with
// its note.
import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { EditorState, Selection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import type { Node as ProseMirrorNode } from "prosemirror-model";
import {
  documentSchema,
  findNodeById,
  footnoteEditorDoc,
  footnoteEditorPlugins,
  footnoteEditorText,
  setFootnoteText,
  type FormulaHost,
  type MathCommandProvider,
} from "../../shared/visualtex-document/src/index.ts";
import { CommandMenu } from "./CommandMenu";

type Lang = "en" | "zh";

/** The node at `pos`, or null where the document no longer reaches. */
const nodeAtSafe = (doc: ProseMirrorNode, pos: number) => (pos >= 0 && pos < doc.content.size ? doc.nodeAt(pos) : null);
const t = (lang: Lang, en: string, zh: string) => (lang === "en" ? en : zh);

export function FootnoteDialog(props: {
  view: EditorView;
  pos: number;
  /** the footnote's id when it was asked for (`pos` may have moved since: the Markdown shown around the caret hidden) */
  id: string | null;
  rect: DOMRect | null;
  lang: Lang;
  host: FormulaHost;
  mathCommands: MathCommandProvider | null;
  /** the document's body font (the note is in it at the foot of the page) */
  font: string;
  onClose: () => void;
}) {
  const { view, pos, lang, onClose } = props;
  // the footnote is the node at `pos` when the dialog opens; it is found again by its id when the text is applied
  const [target] = useState(() => {
    const at = nodeAtSafe(view.state.doc, pos);
    const node = at?.type.name === "footnote" ? at : props.id ? findNodeById(view.state.doc, props.id)?.node : null;
    return node?.type.name === "footnote" ? { id: node.attrs.id as string | null, text: node.attrs.text as string } : null;
  });
  const panelRef = useRef<HTMLFormElement>(null);
  const inputRef = useRef<HTMLDivElement>(null);
  const noteRef = useRef<EditorView | null>(null);
  const [noteState, setNoteState] = useState<EditorState | null>(null);
  const [place, setPlace] = useState<{ left: number; top: number } | null>(null);
  const actions = useRef({ apply: () => {}, close: () => {} });

  // the note's own editor: its text with formulas, the document's formula fields and commands
  useEffect(() => {
    const mount = inputRef.current;
    if (!mount || !target) return;
    const note: EditorView = new EditorView(mount, {
      state: EditorState.create({
        doc: footnoteEditorDoc(documentSchema, target.text),
        plugins: footnoteEditorPlugins(documentSchema, {
          host: props.host,
          mathCommands: props.mathCommands,
          inputMode: () => "markdown",
          submit: () => {
            actions.current.apply();
            return true;
          },
          cancel: () => {
            actions.current.close();
            return true;
          },
        }),
      }),
      attributes: { class: "vtx-document vtx-footnote-editor", spellcheck: "false" },
      dispatchTransaction(tr) {
        const next = note.state.apply(tr);
        note.updateState(next);
        setNoteState(next);
      },
    });
    noteRef.current = note;
    setNoteState(note.state);
    return () => {
      noteRef.current = null;
      note.destroy();
    };
    // the editor is made once, for the footnote the panel opened on
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

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
  }, [view, pos, props.rect]);
  // the caret at the end of the note once the panel is in place
  useEffect(() => {
    const note = noteRef.current;
    if (!place || !note) return;
    note.focus();
    note.dispatch(note.state.tr.setSelection(Selection.atEnd(note.state.doc)));
  }, [place !== null]);

  /** The footnote's position now (the document may have changed around it). */
  const currentPos = (): number | null => {
    if (!target) return null;
    const at = nodeAtSafe(view.state.doc, pos);
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
    const note = noteRef.current;
    if (at !== null && note) setFootnoteText(at, footnoteEditorText(note.state.doc))(view.state, view.dispatch);
    close();
  };
  actions.current = { apply, close };
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
      >
        <div className="vtx-footnote-dialog-title">{t(lang, "Footnote", "脚注")}</div>
        <div ref={inputRef} className="vtx-footnote-input" style={{ ["--vtx-body-font" as string]: props.font }} />
        <CommandMenu view={noteRef.current} state={noteState} lang={lang} frame={panelRef.current} />
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
