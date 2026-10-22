import { useEffect, useState } from "react";
import { MousePointerClick, X } from "lucide-react";
import { useEditorStore } from "../stores/editorStore";
import { useFormulaHotkeyStore } from "../stores/formulaHotkeyStore";
import { useFormulaRowHotkeyStore } from "../stores/formulaRowHotkeyStore";
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
  }, [capturing, isEn, shortcuts, formulaBindings, rowBindings, removeFormulaBinding, removeRowBinding, setButtonShortcut]);

  const groups: { title: string; scope: "both" | "main" | "office" }[] = [
    { title: isEn ? "Main window and editor window" : "主程序与编辑器窗口", scope: "both" },
    { title: isEn ? "Main window" : "主程序", scope: "main" },
    { title: isEn ? "Office editor window" : "Office 编辑器窗口", scope: "office" },
  ];

  return (
    <section className="shortcut-settings-group" data-button-shortcuts>
      <div className="settings-section-title">
        <MousePointerClick size={18} />
        <div>
          <h3>{isEn ? "Buttons" : "常用按钮"}</h3>
        </div>
      </div>
      <p className="office-formula-shortcut-help">
        {isEn
          ? "Give the buttons you use often a shortcut. None is set at first; click a shortcut box and press the keys (Esc cancels)."
          : "为常用按钮设置快捷键。初始均为“无”；点击快捷键框后按下组合键即可（Esc 取消）。"}
      </p>
      {message && (
        <p className="office-formula-shortcut-help" data-button-shortcut-message role="status">
          {message}
        </p>
      )}
      {groups.map((group) => (
        <section className="office-formula-shortcut-host" key={group.scope} data-button-shortcut-scope={group.scope}>
          <header className="office-formula-shortcut-host-header">
            <strong>{group.title}</strong>
          </header>
          <div className="office-formula-shortcut-list">
            {BUTTON_SHORTCUT_ACTIONS.filter((a) => a.scope === group.scope).map((action) => {
              const chord = shortcuts[action.id];
              const isCapturing = capturing === action.id;
              const label = isEn ? action.labelEn : action.labelZh;
              return (
                <div className="office-formula-shortcut-row is-format-fixed" data-button-shortcut={action.id} key={action.id}>
                  <div className="office-formula-shortcut-copy">
                    <strong>{label}</strong>
                  </div>
                  <button
                    type="button"
                    className={`office-formula-shortcut-capture${isCapturing ? " is-capturing" : ""}`}
                    aria-pressed={isCapturing}
                    data-button-shortcut-capture={action.id}
                    {...(isCapturing ? { "data-button-shortcut-capturing": "" } : {})}
                    onClick={() => {
                      setMessage("");
                      setCapturing(isCapturing ? null : action.id);
                    }}
                  >
                    <kbd>{isCapturing ? (isEn ? "Press keys…" : "请按键…") : chord ? formatFormulaHotkeyChord(chord) : isEn ? "None" : "无"}</kbd>
                  </button>
                  <button
                    type="button"
                    className="office-formula-shortcut-clear"
                    disabled={!chord && !isCapturing}
                    data-button-shortcut-clear={action.id}
                    onClick={() => {
                      setButtonShortcut(action.id, null);
                      if (isCapturing) setCapturing(null);
                      setMessage("");
                    }}
                    aria-label={isEn ? `Clear the shortcut of ${label}` : `清除“${label}”的快捷键`}
                    title={isEn ? "Set to None" : "设为无"}
                  >
                    <X size={14} />
                    {isEn ? "None" : "无"}
                  </button>
                </div>
              );
            })}
          </div>
        </section>
      ))}
    </section>
  );
}
