// The document's right-click menu (Mathcha's): copy variants, cut, delete; in a table, rows and columns. The
// formatting controls are DocumentControls.tsx (the workspace's bottom bar), the "\" menu CommandMenu.tsx.
import type { Command, EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { CellSelection } from "prosemirror-tables";
import { isInTable, tableCommands } from "../../shared/visualtex-document/src/index.ts";

export type Lang = "en" | "zh";
const t = (lang: Lang, en: string, zh: string) => (lang === "en" ? en : zh);

function run(view: EditorView | null, command: Command) {
  if (!view) return false;
  const done = command(view.state, view.dispatch, view);
  view.focus();
  return done;
}

// The menu never takes the focus: the editor keeps its selection.
const keepFocus = (event: { preventDefault: () => void }) => event.preventDefault();

// ----------------------------------------------------------- context menu --

export interface ContextMenuState {
  x: number;
  y: number;
}

export interface ContextMenuActions {
  copy(): void;
  copyPlainText(): void;
  cut(): void;
  remove(): void;
}

/** Mathcha's right-click menu: copy variants, cut, delete; in a table, rows and columns. */
export function ContextMenu(props: {
  view: EditorView | null;
  state: EditorState | null;
  lang: Lang;
  at: ContextMenuState | null;
  actions: ContextMenuActions;
  onClose: () => void;
}) {
  const { view, state, lang, at, actions, onClose } = props;
  if (!at || !view || !state) return null;
  const inTable = isInTable(state);
  const c = tableCommands;
  const cellSelection = state.selection instanceof CellSelection;
  const empty = state.selection.empty;
  const item = (label: string, fn: () => void, options: { disabled?: boolean; hint?: string } = {}) => (
    <button
      type="button"
      className="mc-menu-item"
      disabled={options.disabled}
      onClick={() => {
        onClose();
        fn();
        view.focus();
      }}
    >
      <span>{label}</span>
      {options.hint && <kbd>{options.hint}</kbd>}
    </button>
  );
  const command = (cmd: Command) => () => run(view, cmd);
  const style = { left: Math.min(at.x, window.innerWidth - 240), top: Math.min(at.y, window.innerHeight - (inTable ? 470 : 230)) };
  return (
    <>
      <div className="mc-dismiss" onMouseDown={onClose} onContextMenu={(e) => { e.preventDefault(); onClose(); }} />
      <div className="mc-panel mc-context" style={style} role="menu" onMouseDown={keepFocus}>
        {item(t(lang, "Copy", "复制"), actions.copy, { disabled: empty, hint: "Ctrl+C" })}
        {item(t(lang, "Copy Plain Text", "复制纯文本"), actions.copyPlainText, { disabled: empty })}
        {item(t(lang, "Cut", "剪切"), actions.cut, { disabled: empty, hint: "Ctrl+X" })}
        {item(t(lang, "Delete", "删除"), actions.remove, { disabled: empty })}
        {inTable && (
          <>
            <div className="mc-menu-sep" />
            {item(t(lang, "Insert Row Above", "在上方插入行"), command(c.addRowBefore))}
            {item(t(lang, "Insert Row Below", "在下方插入行"), command(c.addRowAfter))}
            {item(t(lang, "Insert Column on Left", "在左侧插入列"), command(c.addColumnBefore))}
            {item(t(lang, "Insert Column on Right", "在右侧插入列"), command(c.addColumnAfter))}
            {cellSelection && item(t(lang, "Merge Cells", "合并单元格"), command(c.mergeCells), { disabled: !c.mergeCells(state) })}
            {c.splitCell(state) && item(t(lang, "Unmerge Cells", "拆分单元格"), command(c.splitCell))}
            <div className="mc-menu-sep" />
            {item(t(lang, "Delete Row", "删除行"), command(c.deleteRow))}
            {item(t(lang, "Delete Column", "删除列"), command(c.deleteColumn))}
            {item(t(lang, "Delete Table", "删除表格"), command(c.deleteTable))}
          </>
        )}
      </div>
    </>
  );
}
