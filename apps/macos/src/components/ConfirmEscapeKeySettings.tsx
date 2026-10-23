import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import {
  AlertTriangle,
  Check,
  CornerDownLeft,
  Keyboard,
  Pencil,
  RotateCcw,
  Trash2,
  X,
} from "lucide-react";
import {
  formatFormulaHotkeyChord,
  formulaHotkeyChordFromEvent,
  formulaHotkeyChordId,
  formulaHotkeyHasModifier,
  formulaHotkeyTargetLabel,
  type FormulaHotkeyChord,
} from "../shortcuts/formulaHotkeys";
import {
  confirmEscapeChordProblem,
  DEFAULT_CONFIRM_ESCAPE_CHORD,
  sameConfirmEscapeChord,
} from "../shortcuts/confirmEscapeKey";
import {
  formulaRowHotkeyAction,
  formulaRowHotkeyActionLabel,
  formulaRowHotkeyChordId,
} from "../shortcuts/formulaRowHotkeys";
import {
  buttonShortcutLabel,
  buttonShortcutWithChord,
} from "../shortcuts/buttonShortcuts";
import { useConfirmEscapeKeyStore } from "../stores/confirmEscapeKeyStore";
import { useEditorStore } from "../stores/editorStore";
import { useFormulaHotkeyStore } from "../stores/formulaHotkeyStore";
import { useFormulaRowHotkeyStore } from "../stores/formulaRowHotkeyStore";

function copy(isEn: boolean) {
  return {
    section: isEn ? "Input" : "输入",
    label: isEn ? "Confirm escape" : "确认转义",
    description: isEn
      ? "Turns only the name right before the caret into its symbol (alpha → α); Backspace right after undoes it."
      : "只把光标前的命令名转成符号（alpha → α）；紧接着按 Backspace 可撤销。",
  };
}

export function ConfirmEscapeKeySettingsSection() {
  const language = useEditorStore((state) => state.language);
  const chord = useConfirmEscapeKeyStore((state) => state.chord);
  const setChord = useConfirmEscapeKeyStore((state) => state.setChord);
  const reset = useConfirmEscapeKeyStore((state) => state.reset);
  const [recording, setRecording] = useState(false);
  const closeRecorder = useCallback(() => setRecording(false), []);
  const isEn = language === "en";
  const text = copy(isEn);
  const isDefault = sameConfirmEscapeChord(chord, DEFAULT_CONFIRM_ESCAPE_CHORD);

  return (
    <section className="office-hotkey-group formula-row-hotkey-group" data-confirm-escape-key>
      <header className="formula-hotkey-section-heading">
        <CornerDownLeft size={16} />
        <strong>{text.section}</strong>
        <span>{isEn ? "Typing in the formula editor" : "在公式编辑区中输入"}</span>
      </header>
      <div className="office-hotkey-action-list">
        <article className="office-hotkey-action-row" data-confirm-escape-key-row>
          <div className="office-hotkey-action-icon" aria-hidden="true">
            <CornerDownLeft size={18} />
          </div>
          <div className="formula-hotkey-binding-copy">
            <strong>{text.label}</strong>
            <span>{text.description}</span>
          </div>
          {chord ? (
            <kbd>{formatFormulaHotkeyChord(chord)}</kbd>
          ) : (
            <span className="office-hotkey-unassigned">{isEn ? "Off" : "已关闭"}</span>
          )}
          <div className="formula-hotkey-binding-actions">
            <button
              type="button"
              className="icon-button compact"
              onClick={() => setRecording(true)}
              aria-label={isEn ? `Change key for ${text.label}` : `修改${text.label}的按键`}
              title={isEn ? "Change key" : "修改按键"}
            >
              <Pencil size={14} />
            </button>
            {!isDefault && (
              <button
                type="button"
                className="icon-button compact"
                onClick={reset}
                aria-label={isEn ? `Restore the default key for ${text.label}` : `恢复${text.label}的默认按键`}
                title={isEn ? "Restore default" : "恢复默认"}
              >
                <RotateCcw size={14} />
              </button>
            )}
            {chord && (
              <button
                type="button"
                className="icon-button compact is-danger"
                onClick={() => setChord(null)}
                aria-label={isEn ? `Turn off ${text.label}` : `关闭${text.label}`}
                title={isEn ? "Turn off" : "关闭"}
              >
                <Trash2 size={14} />
              </button>
            )}
          </div>
        </article>
      </div>
      {recording &&
        createPortal(
          <ConfirmEscapeKeyRecorderDialog onClose={closeRecorder} />,
          document.body,
        )}
    </section>
  );
}

function ConfirmEscapeKeyRecorderDialog({ onClose }: { onClose: () => void }) {
  const dialogRef = useRef<HTMLElement>(null);
  const language = useEditorStore((state) => state.language);
  const buttonShortcuts = useEditorStore((state) => state.buttonShortcuts);
  const setButtonShortcut = useEditorStore((state) => state.setButtonShortcut);
  const existingChord = useConfirmEscapeKeyStore((state) => state.chord);
  const setChord = useConfirmEscapeKeyStore((state) => state.setChord);
  const rowBindings = useFormulaRowHotkeyStore((state) => state.bindings);
  const removeRowBinding = useFormulaRowHotkeyStore((state) => state.removeBinding);
  const formulaBindings = useFormulaHotkeyStore((state) => state.bindings);
  const removeFormulaBinding = useFormulaHotkeyStore((state) => state.removeBinding);
  const [capturedChord, setCapturedChord] = useState<FormulaHotkeyChord | null>(
    existingChord,
  );
  const isEn = language === "en";
  const text = copy(isEn);

  useEffect(() => {
    const frame = window.requestAnimationFrame(() => dialogRef.current?.focus());
    const handleKeyDown = (event: KeyboardEvent) => {
      if (
        event.key === "Escape" &&
        !event.ctrlKey &&
        !event.altKey &&
        !event.shiftKey &&
        !event.metaKey
      ) {
        event.preventDefault();
        event.stopImmediatePropagation();
        onClose();
        return;
      }
      event.preventDefault();
      event.stopImmediatePropagation();
      if (event.repeat || event.isComposing || event.key === "Process") return;
      if (event.getModifierState("AltGraph")) return;
      const chord = formulaHotkeyChordFromEvent(event);
      if (chord) setCapturedChord(chord);
    };
    document.addEventListener("keydown", handleKeyDown, true);
    return () => {
      window.cancelAnimationFrame(frame);
      document.removeEventListener("keydown", handleKeyDown, true);
    };
  }, [onClose]);

  // Chords with a modifier are matched by the editor's hotkeys first.
  const conflicts = useMemo(() => {
    if (!capturedChord || !formulaHotkeyHasModifier(capturedChord)) {
      return { row: null, formula: null, button: null };
    }
    const rowChordId = formulaRowHotkeyChordId(capturedChord);
    const chordId = formulaHotkeyChordId(capturedChord);
    return {
      row:
        rowBindings.find(
          (binding) =>
            binding.chord !== null &&
            formulaRowHotkeyChordId(binding.chord) === rowChordId,
        ) ?? null,
      formula:
        formulaBindings.find(
          (binding) => formulaHotkeyChordId(binding.chord) === chordId,
        ) ?? null,
      button: buttonShortcutWithChord(buttonShortcuts, capturedChord),
    };
  }, [buttonShortcuts, capturedChord, formulaBindings, rowBindings]);

  const problem = capturedChord
    ? confirmEscapeChordProblem(capturedChord, language)
    : null;
  const canSave = Boolean(capturedChord && !problem);
  const conflictLabel = conflicts.row
    ? formulaRowHotkeyActionLabel(formulaRowHotkeyAction(conflicts.row.actionId), language)
    : conflicts.formula
      ? formulaHotkeyTargetLabel(conflicts.formula.target, language)
      : conflicts.button
        ? buttonShortcutLabel(conflicts.button, language)
        : null;

  const save = () => {
    if (!capturedChord || !canSave) return;
    if (conflicts.button) setButtonShortcut(conflicts.button, null);
    if (conflicts.row) removeRowBinding(conflicts.row.actionId);
    if (conflicts.formula) removeFormulaBinding(conflicts.formula.id);
    setChord(capturedChord);
    onClose();
  };

  return (
    <div
      className="modal-backdrop formula-hotkey-modal-backdrop"
      role="presentation"
      onMouseDown={onClose}
    >
      <section
        ref={dialogRef}
        className="formula-hotkey-recorder-dialog formula-row-hotkey-recorder-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="confirm-escape-key-recorder-title"
        tabIndex={-1}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="dialog-header">
          <div>
            <span className="eyebrow">ESCAPE KEY</span>
            <h2 id="confirm-escape-key-recorder-title">
              {isEn ? "Set confirm-escape key" : "设置确认转义键"}
            </h2>
          </div>
          <button
            type="button"
            className="icon-button"
            onClick={onClose}
            aria-label={isEn ? "Close" : "关闭"}
          >
            <X size={18} />
          </button>
        </header>

        <div className="formula-hotkey-recorder-content">
          <div className="formula-hotkey-target-card office-hotkey-target-card">
            <div className="formula-hotkey-target-preview">
              <CornerDownLeft size={28} />
            </div>
            <div>
              <strong>{text.label}</strong>
              <span>{text.description}</span>
              <code>
                {isEn ? "Default: " : "默认："}
                {formatFormulaHotkeyChord(DEFAULT_CONFIRM_ESCAPE_CHORD)}
              </code>
            </div>
          </div>

          <div className="formula-hotkey-capture-box" aria-live="polite">
            <Keyboard size={20} />
            {capturedChord ? (
              <kbd>{formatFormulaHotkeyChord(capturedChord)}</kbd>
            ) : (
              <strong>{isEn ? "Press a key now" : "现在按下需要使用的按键"}</strong>
            )}
            <span>
              {isEn
                ? "A plain key such as Space works. When nothing can be escaped, the key does its usual job."
                : "可以是 Space 这样的单键；光标前没有可转义的命令时，该键照常起作用。"}
            </span>
          </div>

          {problem && (
            <div className="formula-hotkey-message is-danger" role="alert">
              <AlertTriangle size={16} />
              <span>{problem}</span>
            </div>
          )}
          {conflictLabel && !problem && (
            <div className="formula-hotkey-message is-warning" role="alert">
              <AlertTriangle size={16} />
              <span>
                {isEn
                  ? `Currently assigned to “${conflictLabel}”. Saving will replace that assignment.`
                  : `当前已分配给“${conflictLabel}”，保存后将替换原绑定。`}
              </span>
            </div>
          )}
          {capturedChord && canSave && !conflictLabel && (
            <div className="formula-hotkey-message is-success">
              <Check size={16} />
              <span>{isEn ? "Available." : "可以使用。"}</span>
            </div>
          )}
        </div>

        <footer className="dialog-footer formula-hotkey-recorder-footer">
          <span>
            {existingChord
              ? isEn
                ? `Current: ${formatFormulaHotkeyChord(existingChord)}`
                : `当前：${formatFormulaHotkeyChord(existingChord)}`
              : isEn
                ? "Quick escape is off"
                : "确认转义已关闭"}
          </span>
          <div>
            <button type="button" className="secondary-button" onClick={onClose}>
              {isEn ? "Cancel" : "取消"}
            </button>
            <button
              type="button"
              className="primary-button"
              disabled={!canSave}
              onClick={save}
            >
              {conflictLabel
                ? isEn
                  ? "Replace and save"
                  : "替换并保存"
                : isEn
                  ? "Save"
                  : "保存"}
            </button>
          </div>
        </footer>
      </section>
    </div>
  );
}
