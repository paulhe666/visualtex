import { useEffect, useMemo, useRef, useState } from "react";
import { AlertTriangle, Check, Keyboard, Rows3, X } from "lucide-react";
import {
  formatFormulaHotkeyChord,
  formulaHotkeyChordFromEvent,
  formulaHotkeyChordId,
  formulaHotkeyTargetLabel,
  type FormulaHotkeyChord,
} from "../shortcuts/formulaHotkeys";
import {
  defaultFormulaRowHotkeyChord,
  formulaRowHotkeyAction,
  formulaRowHotkeyActionLabel,
  formulaRowHotkeyChordId,
  formulaRowHotkeyChordProblem,
  type FormulaRowHotkeyActionDefinition,
} from "../shortcuts/formulaRowHotkeys";
import { useEditorStore } from "../stores/editorStore";
import { buttonShortcutLabel, buttonShortcutWithChord } from "../shortcuts/buttonShortcuts";
import { useFormulaHotkeyStore } from "../stores/formulaHotkeyStore";
import { useFormulaRowHotkeyStore } from "../stores/formulaRowHotkeyStore";

interface Props {
  action: FormulaRowHotkeyActionDefinition | null;
  onClose: () => void;
}

export function FormulaRowHotkeyRecorderDialog({ action, onClose }: Props) {
  const dialogRef = useRef<HTMLElement>(null);
  const language = useEditorStore((state) => state.language);
  const buttonShortcuts = useEditorStore((state) => state.buttonShortcuts);
  const setButtonShortcut = useEditorStore((state) => state.setButtonShortcut);
  const rowBindings = useFormulaRowHotkeyStore((state) => state.bindings);
  const setRowBinding = useFormulaRowHotkeyStore((state) => state.setBinding);
  const formulaBindings = useFormulaHotkeyStore((state) => state.bindings);
  const removeFormulaBinding = useFormulaHotkeyStore(
    (state) => state.removeBinding,
  );
  const [capturedChord, setCapturedChord] = useState<FormulaHotkeyChord | null>(
    null,
  );
  const isEn = language === "en";
  const existingChord = action
    ? rowBindings.find((binding) => binding.actionId === action.id)?.chord ??
      null
    : null;

  useEffect(() => {
    setCapturedChord(existingChord);
  }, [action?.id, existingChord]);

  useEffect(() => {
    if (!action) return;
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
  }, [action, onClose]);

  const conflicts = useMemo(() => {
    if (!action || !capturedChord) return { row: null, formula: null };
    const rowChordId = formulaRowHotkeyChordId(capturedChord);
    const chordId = formulaHotkeyChordId(capturedChord);
    return {
      row:
        rowBindings.find(
          (binding) =>
            binding.actionId !== action.id &&
            binding.chord !== null &&
            formulaRowHotkeyChordId(binding.chord) === rowChordId,
        ) ?? null,
      formula:
        formulaBindings.find(
          (binding) => formulaHotkeyChordId(binding.chord) === chordId,
        ) ?? null,
    };
  }, [action, capturedChord, formulaBindings, rowBindings]);

  if (!action) return null;
  const buttonConflict = capturedChord ? buttonShortcutWithChord(buttonShortcuts, capturedChord) : null;

  const problem = capturedChord
    ? formulaRowHotkeyChordProblem(capturedChord, language)
    : null;
  const canSave = Boolean(capturedChord && !problem);
  const actionLabel = formulaRowHotkeyActionLabel(action, language);
  const conflictLabel = conflicts.row
    ? formulaRowHotkeyActionLabel(
        formulaRowHotkeyAction(conflicts.row.actionId),
        language,
      )
    : conflicts.formula
      ? formulaHotkeyTargetLabel(conflicts.formula.target, language)
      : buttonConflict
        ? buttonShortcutLabel(buttonConflict, language)
        : null;

  const saveBinding = () => {
    if (!capturedChord || !canSave) return;
    // a button shortcut is matched first: it gives the chord up
    if (buttonConflict) setButtonShortcut(buttonConflict, null);
    if (conflicts.formula) removeFormulaBinding(conflicts.formula.id);
    setRowBinding(action.id, capturedChord);
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
        aria-labelledby="formula-row-hotkey-recorder-title"
        tabIndex={-1}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="dialog-header">
          <div>
            <span className="eyebrow">ROW HOT KEY</span>
            <h2 id="formula-row-hotkey-recorder-title">
              {isEn ? "Set formula row hotkey" : "设置公式行快捷键"}
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
              <Rows3 size={28} />
            </div>
            <div>
              <strong>{actionLabel}</strong>
              <span>{isEn ? action.descriptionEn : action.descriptionZh}</span>
              <code>
                {isEn ? "Default: " : "默认："}
                {formatFormulaHotkeyChord(defaultFormulaRowHotkeyChord(action.id))}
              </code>
            </div>
          </div>

          <div className="formula-hotkey-capture-box" aria-live="polite">
            <Keyboard size={20} />
            {capturedChord ? (
              <kbd>{formatFormulaHotkeyChord(capturedChord)}</kbd>
            ) : (
              <strong>
                {isEn ? "Press a shortcut now" : "现在按下需要绑定的快捷键"}
              </strong>
            )}
            <span>
              {isEn
                ? "Works while a formula row has focus in the editor."
                : "在编辑区的公式行中生效。"}
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
                ? "No hotkey assigned"
                : "尚未设置快捷键"}
          </span>
          <div>
            <button type="button" className="secondary-button" onClick={onClose}>
              {isEn ? "Cancel" : "取消"}
            </button>
            <button
              type="button"
              className="primary-button"
              disabled={!canSave}
              onClick={saveBinding}
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
