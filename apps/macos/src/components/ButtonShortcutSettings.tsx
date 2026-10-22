import { useEffect, useState } from "react";
import { MousePointerClick, Pencil, Plus, Trash2 } from "lucide-react";
import { useEditorStore } from "../stores/editorStore";
import { useFormulaHotkeyStore } from "../stores/formulaHotkeyStore";
import { useFormulaRowHotkeyStore } from "../stores/formulaRowHotkeyStore";
import { useOfficeHotkeyStore } from "../stores/officeHotkeyStore";
import { officeHotkeyAction, officeHotkeyActionLabel } from "../shortcuts/officeHotkeys";
import {
  BUTTON_SHORTCUT_ACTIONS,
  buttonShortcutChordProblem,
  buttonShortcutWithChord,
  type ButtonShortcutId,
} from "../shortcuts/buttonShortcuts";
import {
  formatFormulaHotkeyChord,
  formulaHotkeyChordFromEvent,
  formulaHotkeyChordId,
  formulaHotkeyTargetLabel,
} from "../shortcuts/formulaHotkeys";
import { FORMULA_ROW_HOTKEY_ACTIONS } from "../shortcuts/formulaRowHotkeys";

/**
 * Settings > Keyboard shortcuts > Buttons: a chord for each common button of
 * the main window and the Office editor window; none at first. A chord a
 * formula or row shortcut had moves here (button shortcuts are matched first).
 */
export function ButtonShortcutSettingsSection() {
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const shortcuts = useEditorStore((state) => state.buttonShortcuts);
  const setButtonShortcut = useEditorStore((state) => state.setButtonShortcut);
  const formulaBindings = useFormulaHotkeyStore((state) => state.bindings);
  const removeFormulaBinding = useFormulaHotkeyStore((state) => state.removeBinding);
  const rowBindings = useFormulaRowHotkeyStore((state) => state.bindings);
  const removeRowBinding = useFormulaRowHotkeyStore((state) => state.removeBinding);
  // Office global hotkeys are system-wide (registered by the app): change them in their own section
  const officeBindings = useOfficeHotkeyStore((state) => state.bindings);
  const [capturing, setCapturing] = useState<ButtonShortcutId | null>(null);
  const [message, setMessage] = useState("");
  const labelOf = (id: ButtonShortcutId) => {
    const action = BUTTON_SHORTCUT_ACTIONS.find((a) => a.id === id)!;
    return isEn ? action.labelEn : action.labelZh;
  };

  useEffect(() => {
    if (!capturing) return;
    const onKeyDown = (event: KeyboardEvent) => {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (event.repeat || event.isComposing || event.key === "Process") return;
      if (event.key === "Escape" && !event.ctrlKey && !event.altKey && !event.shiftKey && !event.metaKey) {
        setCapturing(null);
        setMessage("");
        return;
      }
      if (event.getModifierState("AltGraph")) {
        setMessage(isEn ? "AltGr is used for text input and cannot be assigned." : "AltGr 用于文字输入，不能设置为快捷键。");
        return;
      }
      const chord = formulaHotkeyChordFromEvent(event);
      if (!chord) return; // a modifier alone: wait for the key
      const problem = buttonShortcutChordProblem(chord, isEn ? "en" : "cn");
      if (problem) {
        setMessage(problem);
        return;
      }
      const key = formulaHotkeyChordId(chord);
      const officeBinding = officeBindings.find((binding) => formulaHotkeyChordId(binding.chord) === key);
      if (officeBinding) {
        const label = officeHotkeyActionLabel(officeHotkeyAction(officeBinding.actionId), isEn ? "en" : "cn");
        setMessage(
          isEn
            ? `This shortcut is the Office global shortcut “${label}”. Change that one first.`
            : `该快捷键已用作 Office 全局快捷键“${label}”，请先修改那一项。`,
        );
        return;
      }
      const notes: string[] = [];
      const otherButton = buttonShortcutWithChord(shortcuts, chord);
      if (otherButton && otherButton !== capturing) {
        notes.push(isEn ? `moved from “${labelOf(otherButton)}”` : `已从“${labelOf(otherButton)}”移到这里`);
      }
      for (const binding of formulaBindings) {
        if (formulaHotkeyChordId(binding.chord) !== key) continue;
        removeFormulaBinding(binding.id);
        const label = formulaHotkeyTargetLabel(binding.target, isEn ? "en" : "cn");
        notes.push(isEn ? `removed from the formula shortcut “${label}”` : `已从公式快捷键“${label}”移除`);
      }
      for (const binding of rowBindings) {
        if (!binding.chord || formulaHotkeyChordId(binding.chord) !== key) continue;
        removeRowBinding(binding.actionId);
        const action = FORMULA_ROW_HOTKEY_ACTIONS.find((a) => a.id === binding.actionId);
        const label = action ? (isEn ? action.labelEn : action.labelZh) : binding.actionId;
        notes.push(isEn ? `removed from the row shortcut “${label}”` : `已从公式行快捷键“${label}”移除`);
      }
      setButtonShortcut(capturing, chord);
      setCapturing(null);
      setMessage(notes.length ? `${formatFormulaHotkeyChord(chord)}: ${notes.join(isEn ? "; " : "；")}` : "");
    };
    document.addEventListener("keydown", onKeyDown, true);
    return () => document.removeEventListener("keydown", onKeyDown, true);
  }, [capturing, isEn, shortcuts, formulaBindings, rowBindings, officeBindings, removeFormulaBinding, removeRowBinding, setButtonShortcut]);

  const groups: { title: string; scope: "both" | "main" | "office" }[] = [
    { title: isEn ? "Main window and editor window" : "主程序与编辑器窗口", scope: "both" },
    { title: isEn ? "Main window" : "主程序", scope: "main" },
    { title: isEn ? "Office editor window" : "Office 编辑器窗口", scope: "office" },
  ];

  return (
    <section className="office-hotkey-group" data-button-shortcuts>
      <header className="formula-hotkey-section-heading">
        <MousePointerClick size={16} />
        <strong>{isEn ? "Buttons" : "常用按钮"}</strong>
        <span>
          {isEn
            ? "None at first. Click + and press the keys (Esc cancels)."
            : "初始均为“未设置”。点击 + 后按下组合键（Esc 取消）。"}
        </span>
      </header>
      {message && (
        <p className="formula-hotkey-recorder-note" data-button-shortcut-message role="status">
          {message}
        </p>
      )}
      {groups.map((group) => (
        <div className="office-hotkey-action-list" key={group.scope} data-button-shortcut-scope={group.scope}>
          {BUTTON_SHORTCUT_ACTIONS.filter((a) => a.scope === group.scope).map((action) => {
            const chord = shortcuts[action.id];
            const isCapturing = capturing === action.id;
            const label = isEn ? action.labelEn : action.labelZh;
            return (
              <article className="office-hotkey-action-row" data-button-shortcut={action.id} key={action.id}>
                <div className="office-hotkey-action-icon" aria-hidden="true">
                  <MousePointerClick size={18} />
                </div>
                <div className="formula-hotkey-binding-copy">
                  <strong>{label}</strong>
                  <span>{group.title}</span>
                </div>
                {isCapturing ? (
                  <kbd data-button-shortcut-capturing="">{isEn ? "Press keys…" : "请按键…"}</kbd>
                ) : chord ? (
                  <kbd>{formatFormulaHotkeyChord(chord)}</kbd>
                ) : (
                  <span className="office-hotkey-unassigned">{isEn ? "Not set" : "未设置"}</span>
                )}
                <div className="formula-hotkey-binding-actions">
                  <button
                    type="button"
                    className="icon-button compact"
                    aria-pressed={isCapturing}
                    data-button-shortcut-capture={action.id}
                    onClick={() => {
                      setMessage("");
                      setCapturing(isCapturing ? null : action.id);
                    }}
                    aria-label={
                      chord
                        ? isEn ? `Change the shortcut of ${label}` : `修改“${label}”的快捷键`
                        : isEn ? `Set a shortcut for ${label}` : `设置“${label}”的快捷键`
                    }
                    title={chord ? (isEn ? "Change shortcut" : "修改快捷键") : (isEn ? "Set shortcut" : "设置快捷键")}
                  >
                    {chord ? <Pencil size={14} /> : <Plus size={14} />}
                  </button>
                  {(chord || isCapturing) && (
                    <button
                      type="button"
                      className="icon-button compact is-danger"
                      data-button-shortcut-clear={action.id}
                      onClick={() => {
                        setButtonShortcut(action.id, null);
                        if (isCapturing) setCapturing(null);
                        setMessage("");
                      }}
                      aria-label={isEn ? `Clear the shortcut of ${label}` : `清除“${label}”的快捷键`}
                      title={isEn ? "Set to none" : "设为未设置"}
                    >
                      <Trash2 size={14} />
                    </button>
                  )}
                </div>
              </article>
            );
          })}
        </div>
      ))}
    </section>
  );
}
