import { useEffect, useState } from "react";
import { Pencil, Plus, RotateCcw, Sigma, Trash2 } from "lucide-react";
import { useEditorStore } from "../stores/editorStore";
import { useOfficeHotkeyStore } from "../stores/officeHotkeyStore";
import { configureAnywhereGlobalHotkey, useAnywhereHotkeyStore } from "../stores/anywhereHotkeyStore";
import { officeHotkeyAction, officeHotkeyActionLabel } from "../shortcuts/officeHotkeys";
import {
  ANYWHERE_HOTKEY_ACTIONS,
  anywhereHotkeyChordProblem,
  anywhereHotkeyDefinition,
  type AnywhereHotkeyAction,
} from "../shortcuts/anywhereHotkeys";
import {
  formatFormulaHotkeyChord,
  formulaHotkeyChordFromEvent,
  formulaHotkeyChordId,
  type FormulaHotkeyChord,
} from "../shortcuts/formulaHotkeys";

/**
 * Settings > Keyboard shortcuts > Formula input method (VisualTeX Anywhere):
 * the system-wide shortcut that opens and closes a formula in any
 * application, and the shortcuts used while a formula is open (output
 * format, screenshot OCR). Click the pencil and press the keys (Esc cancels).
 */
export function AnywhereHotkeySettingsSection() {
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const bindings = useAnywhereHotkeyStore((state) => state.bindings);
  const setBinding = useAnywhereHotkeyStore((state) => state.setBinding);
  const officeBindings = useOfficeHotkeyStore((state) => state.bindings);
  const [capturing, setCapturing] = useState<AnywhereHotkeyAction | null>(null);
  const [message, setMessage] = useState("");
  const labelOf = (id: AnywhereHotkeyAction) => {
    const action = anywhereHotkeyDefinition(id);
    return isEn ? action.labelEn : action.labelZh;
  };

  const apply = async (id: AnywhereHotkeyAction, chord: FormulaHotkeyChord | null) => {
    const action = anywhereHotkeyDefinition(id);
    if (action.global) {
      try {
        await configureAnywhereGlobalHotkey(id, chord);
      } catch (error) {
        setMessage(error instanceof Error ? error.message : String(error));
        return false;
      }
    }
    setBinding(id, chord);
    return true;
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
      const chord = formulaHotkeyChordFromEvent(event);
      if (!chord) return; // a modifier alone: wait for the key
      const problem = anywhereHotkeyChordProblem(capturing, chord, isEn ? "en" : "cn");
      if (problem) {
        setMessage(problem);
        return;
      }
      const key = formulaHotkeyChordId(chord);
      if (anywhereHotkeyDefinition(capturing).global) {
        const office = officeBindings.find((binding) => formulaHotkeyChordId(binding.chord) === key);
        if (office) {
          const label = officeHotkeyActionLabel(officeHotkeyAction(office.actionId), isEn ? "en" : "cn");
          setMessage(
            isEn
              ? `Already used by the Office global shortcut “${label}”.`
              : `已被 Office 全局快捷键“${label}”占用。`,
          );
          return;
        }
      }
      const previous = ANYWHERE_HOTKEY_ACTIONS.find(
        (action) => action.id !== capturing && bindings[action.id] && formulaHotkeyChordId(bindings[action.id]!) === key,
      );
      if (previous?.global) {
        setMessage(
          isEn
            ? `Already used by “${labelOf(previous.id)}”.`
            : `已被“${labelOf(previous.id)}”占用。`,
        );
        return;
      }
      const target = capturing;
      void apply(target, chord).then((applied) => {
        if (!applied) return;
        setCapturing(null);
        setMessage(
          previous
            ? `${formatFormulaHotkeyChord(chord)}: ${isEn ? `moved from “${labelOf(previous.id)}”` : `已从“${labelOf(previous.id)}”移到这里`}`
            : "",
        );
      });
    };
    document.addEventListener("keydown", onKeyDown, true);
    return () => document.removeEventListener("keydown", onKeyDown, true);
  }, [capturing, isEn, bindings, officeBindings]);

  return (
    <section className="office-hotkey-group" data-anywhere-shortcuts>
      <header className="formula-hotkey-section-heading">
        <Sigma size={16} />
        <strong>{isEn ? "Formula input method" : "公式输入法"}</strong>
      </header>
      {message && (
        <p className="formula-hotkey-recorder-note" data-anywhere-shortcut-message role="status">
          {message}
        </p>
      )}
      <div className="office-hotkey-action-list">
        {ANYWHERE_HOTKEY_ACTIONS.map((action) => {
          const chord = bindings[action.id];
          const isCapturing = capturing === action.id;
          const label = isEn ? action.labelEn : action.labelZh;
          const isDefault =
            (chord ? formulaHotkeyChordId(chord) : null) ===
            (action.defaultChord ? formulaHotkeyChordId(action.defaultChord) : null);
          return (
            <article className="office-hotkey-action-row" data-anywhere-shortcut={action.id} key={action.id}>
              <div className="office-hotkey-action-icon" aria-hidden="true">
                <Sigma size={18} />
              </div>
              <div className="formula-hotkey-binding-copy">
                <strong>{label}</strong>
                <span>
                  {action.global
                    ? isEn ? "In every application" : "在任意应用中"
                    : isEn ? "While a formula is open" : "公式模式中"}
                </span>
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
                  onClick={() => {
                    setMessage("");
                    setCapturing(isCapturing ? null : action.id);
                  }}
                  aria-label={chord ? (isEn ? `Change the shortcut of ${label}` : `修改“${label}”的快捷键`) : isEn ? `Set a shortcut for ${label}` : `设置“${label}”的快捷键`}
                  title={chord ? (isEn ? "Change shortcut" : "修改快捷键") : isEn ? "Set shortcut" : "设置快捷键"}
                >
                  {chord ? <Pencil size={14} /> : <Plus size={14} />}
                </button>
                {action.defaultChord && !isDefault && (
                  <button
                    type="button"
                    className="icon-button compact"
                    onClick={() => {
                      setCapturing(null);
                      setMessage("");
                      void apply(action.id, action.defaultChord);
                    }}
                    aria-label={isEn ? `Restore the default shortcut of ${label}` : `恢复“${label}”的默认快捷键`}
                    title={isEn ? "Restore default" : "恢复默认"}
                  >
                    <RotateCcw size={14} />
                  </button>
                )}
                {action.id !== "anywhere-toggle" && (chord || isCapturing) && (
                  <button
                    type="button"
                    className="icon-button compact is-danger"
                    onClick={() => {
                      setCapturing(null);
                      setMessage("");
                      void apply(action.id, null);
                    }}
                    aria-label={isEn ? `Clear the shortcut of ${label}` : `清除“${label}”的快捷键`}
                    title={isEn ? "Clear shortcut" : "清除快捷键"}
                  >
                    <Trash2 size={14} />
                  </button>
                )}
              </div>
            </article>
          );
        })}
      </div>
    </section>
  );
}
