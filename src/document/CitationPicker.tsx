// The citation picker (DESIGN.md §8.3): the reference list's entries to cite (one or several), or a new entry to add
// (and cite). In the toolbar's citation button, and at the caret for the "\\cite" / "/cite" menu entries.
import { useState } from "react";
import type { Command, EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { addReference, bibliographyEntries, insertCitation } from "../../shared/visualtex-document/src/index.ts";

type Lang = "en" | "zh";
const t = (lang: Lang, en: string, zh: string) => (lang === "en" ? en : zh);

function run(view: EditorView | null, command: Command) {
  if (!view) return false;
  const done = command(view.state, view.dispatch, view);
  view.focus();
  return done;
}

/** A press in a field: it may take the focus (the panels otherwise keep the editor's). */
const letFocus = (event: { stopPropagation: () => void }) => event.stopPropagation();

/** Cite reference list entries, or add a new one and cite it. */
export function CitationPicker(props: { lang: Lang; view: EditorView | null; state: EditorState; close: () => void }) {
  const { lang, view, state, close } = props;
  const entries = bibliographyEntries(state.doc);
  const [chosen, setChosen] = useState<string[]>([]);
  const [text, setText] = useState("");
  const toggle = (id: string) => setChosen((list) => (list.includes(id) ? list.filter((x) => x !== id) : [...list, id]));
  return (
    <div className="document-citation-picker">
      <div className="document-menu-heading">{t(lang, "Cite references", "引用参考文献")}</div>
      <div className="document-citation-entries">
        {entries.map((entry) => (
          <label key={entry.id} className="document-citation-entry" onMouseDown={letFocus}>
            <input type="checkbox" checked={chosen.includes(entry.id)} onChange={() => toggle(entry.id)} />
            <span className="document-citation-number">[{entry.number}]</span>
            <span className="document-citation-text">{entry.text || t(lang, "(empty)", "（空）")}</span>
          </label>
        ))}
      </div>
      {entries.length > 0 && (
        <button
          type="button"
          className="primary-small-button"
          disabled={!chosen.length}
          onClick={() => {
            close();
            run(view, insertCitation(chosen));
          }}
        >
          {t(lang, "Cite", "插入引用")}
        </button>
      )}
      <div className="document-menu-heading">{t(lang, "New reference", "新建参考文献")}</div>
      <textarea
        className="document-citation-new"
        rows={3}
        value={text}
        onMouseDown={letFocus}
        onChange={(event) => setText(event.target.value)}
        spellCheck={false}
      />
      <div className="document-citation-actions">
        <button
          type="button"
          className="secondary-button"
          disabled={!text.trim()}
          onClick={() => {
            close();
            run(view, addReference(text.trim(), false));
          }}
        >
          {t(lang, "Add", "添加")}
        </button>
        <button
          type="button"
          className="primary-small-button"
          disabled={!text.trim()}
          onClick={() => {
            close();
            run(view, addReference(text.trim(), true));
          }}
        >
          {t(lang, "Add and cite", "添加并引用")}
        </button>
      </div>
    </div>
  );
}

