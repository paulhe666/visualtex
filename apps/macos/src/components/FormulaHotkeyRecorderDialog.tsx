import { useEffect, useMemo, useRef, useState } from "react";
import { isTauri } from "@tauri-apps/api/core";
import { AlertTriangle, Check, Keyboard, X } from "lucide-react";
import { MathPreview } from "./MathPreview";
import {
  formatFormulaHotkeyChord,
  formulaHotkeyChordFromEvent,
  formulaHotkeyChordId,
  formulaHotkeyHasModifier,
  formulaHotkeyTargetKindLabel,
  formulaHotkeyTargetLabel,
  protectedFormulaHotkeyAction,
  type FormulaHotkeyChord,
  type FormulaHotkeyTarget,
} from "../shortcuts/formulaHotkeys";
import { useFormulaHotkeyStore } from "../stores/formulaHotkeyStore";
import { buttonShortcutLabel, buttonShortcutWithChord } from "../shortcuts/buttonShortcuts";
import { useOfficeHotkeyStore } from "../stores/officeHotkeyStore";
import { useEditorStore } from "../stores/editorStore";
import {
  officeHotkeyAction,
  officeHotkeyActionLabel,
} from "../shortcuts/officeHotkeys";
import { configureOfficeHotkeys } from "../runtime/officeHotkeys";
import { useFormulaRowHotkeyStore } from "../stores/formulaRowHotkeyStore";
import {
  formulaRowHotkeyAction,
  formulaRowHotkeyActionLabel,
  formulaRowHotkeyChordId,
} from "../shortcuts/formulaRowHotkeys";

interface Props {
  target: FormulaHotkeyTarget | null;
  onClose: () => void;
}

export function FormulaHotkeyRecorderDialog({ target, onClose }: Props) {
  const dialogRef = useRef<HTMLElement>(null);
  const bindings = useFormulaHotkeyStore((state) => state.bindings);
  const setBinding = useFormulaHotkeyStore((state) => state.setBinding);
  const officeBindings = useOfficeHotkeyStore((state) => state.bindings);
  const removeOfficeBinding = useOfficeHotkeyStore(
    (state) => state.removeBinding,
  );
  const rowBindings = useFormulaRowHotkeyStore((state) => state.bindings);
  const removeRowBinding = useFormulaRowHotkeyStore(
    (state) => state.removeBinding,
  );
  const language = useEditorStore((state) => state.language);
  const buttonShortcuts = useEditorStore((state) => state.buttonShortcuts);
  const setButtonShortcut = useEditorStore((state) => state.setButtonShortcut);
  const isEn = language === "en";
  const existingBinding = target
    ? bindings.find((binding) => binding.target.id === target.id) ?? null
    : null;
  const [capturedChord, setCapturedChord] = useState<FormulaHotkeyChord | null>(
    null,
  );
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState("");

  useEffect(() => {
    setCapturedChord(existingBinding?.chord ?? null);
    setSaveError("");
  }, [target?.id, existingBinding?.id]);

  useEffect(() => {
    if (!target) return;
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
  }, [target, onClose]);

  const conflicts = useMemo(() => {
    if (!target || !capturedChord) return { formula: null, office: null, row: null };
    const chordId = formulaHotkeyChordId(capturedChord);
    const rowChordId = formulaRowHotkeyChordId(capturedChord);
    return {
      // A formula-row shortcut is checked first in the editor, so the same
      // chord can only belong to one of them.
      row:
        rowBindings.find(
          (binding) =>
            binding.chord !== null &&
            formulaRowHotkeyChordId(binding.chord) === rowChordId,
        ) ?? null,
      formula:
        bindings.find(
        (binding) =>
          binding.target.id !== target.id &&
          formulaHotkeyChordId(binding.chord) === chordId,
        ) ?? null,
      office:
        officeBindings.find(
          (binding) => formulaHotkeyChordId(binding.chord) === chordId,
        ) ?? null,
    };
  }, [bindings, capturedChord, officeBindings, rowBindings, target]);

  if (!target) return null;
  const buttonConflict = capturedChord ? buttonShortcutWithChord(buttonShortcuts, capturedChord) : null;

  const protectedAction = capturedChord
    ? protectedFormulaHotkeyAction(capturedChord, language)
    : null;
  const hasModifier = capturedChord
    ? formulaHotkeyHasModifier(capturedChord)
    : false;
  const canSave = Boolean(
    capturedChord && hasModifier && !protectedAction && !saving,
  );
  const targetLabel = formulaHotkeyTargetLabel(target, language);
  const conflictLabel = conflicts.row
    ? formulaRowHotkeyActionLabel(
        formulaRowHotkeyAction(conflicts.row.actionId),
        language,
      )
    : conflicts.formula
    ? formulaHotkeyTargetLabel(conflicts.formula.target, language)
    : conflicts.office
      ? officeHotkeyActionLabel(
          officeHotkeyAction(conflicts.office.actionId),
          language,
        )
      : buttonConflict
        ? buttonShortcutLabel(buttonConflict, language)
        : null;

  const saveBinding = async () => {
    if (!capturedChord || !canSave) return;
    setSaving(true);
    setSaveError("");
    try {
      if (conflicts.office && isTauri()) {
        await configureOfficeHotkeys(
          officeBindings.filter(
            (binding) => binding.actionId !== conflicts.office?.actionId,
          ),
        );
      }
      if (conflicts.office) removeOfficeBinding(conflicts.office.actionId);
      if (conflicts.row) removeRowBinding(conflicts.row.actionId);
      // a button shortcut is matched first: it gives the chord up
      if (buttonConflict) setButtonShortcut(buttonConflict, null);
      setBinding(target, capturedChord);
      onClose();
    } catch (error) {
      setSaveError(
        error instanceof Error
          ? error.message
          : String(
              error ||
                (isEn
                  ? "Unable to replace the Office global hotkey."
                  : "无法替换 Office 全局快捷键。"),
            ),
      );
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="modal-backdrop formula-hotkey-modal-backdrop" role="presentation" onMouseDown={onClose}>
      <section
        ref={dialogRef}
        className="formula-hotkey-recorder-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="formula-hotkey-recorder-title"
        tabIndex={-1}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="dialog-header">
          <div>
            <span className="eyebrow">HOT KEY</span>
            <h2 id="formula-hotkey-recorder-title">
              {isEn ? "Set formula hotkey" : "设置公式快捷键"}
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
          <div className="formula-hotkey-target-card">
            <div className="formula-hotkey-target-preview">
              <MathPreview latex={target.command.previewLatex} fit />
            </div>
            <div>
              <strong>{targetLabel}</strong>
              <span>{formulaHotkeyTargetKindLabel(target, language)}</span>
              <code>{target.command.command}</code>
            </div>
          </div>

          <div className="formula-hotkey-capture-box" aria-live="polite">
            <Keyboard size={20} />
            {capturedChord ? (
              <kbd>{formatFormulaHotkeyChord(capturedChord)}</kbd>
            ) : (
              <strong>
                {isEn ? "Press a shortcut" : "请按快捷键"}
              </strong>
            )}
          </div>

          {capturedChord && !hasModifier && (
            <div className="formula-hotkey-message is-warning" role="alert">
              <AlertTriangle size={16} />
              <span>
                {isEn
                  ? "Requires Ctrl, Alt or Command."
                  : "须包含 Ctrl、Alt 或 Command。"}
              </span>
            </div>
          )}

          {protectedAction && (
            <div className="formula-hotkey-message is-danger" role="alert">
              <AlertTriangle size={16} />
              <span>
                {isEn
                  ? `Reserved for ${protectedAction}.`
                  : `已保留给“${protectedAction}”。`}
              </span>
            </div>
          )}

          {conflictLabel && !protectedAction && hasModifier && (
            <div className="formula-hotkey-message is-warning" role="alert">
              <AlertTriangle size={16} />
              <span>
                {isEn
                  ? `Assigned to “${conflictLabel}”.`
                  : `已分配给“${conflictLabel}”。`}
              </span>
            </div>
          )}

          {capturedChord && canSave && !conflictLabel && (
            <div className="formula-hotkey-message is-success">
              <Check size={16} />
              <span>
                {isEn
                  ? "Available"
                  : "可用"}
              </span>
            </div>
          )}
          {saveError && (
            <div className="formula-hotkey-message is-danger" role="alert">
              <AlertTriangle size={16} />
              <span>{saveError}</span>
            </div>
          )}
        </div>

        <footer className="dialog-footer formula-hotkey-recorder-footer">
          <span>
            {existingBinding
              ? isEn
                ? `Current: ${formatFormulaHotkeyChord(existingBinding.chord)}`
                : `当前：${formatFormulaHotkeyChord(existingBinding.chord)}`
              : isEn
                ? "Not set"
                : "未设置"}
          </span>
          <div>
            <button type="button" className="secondary-button" onClick={onClose}>
              {isEn ? "Cancel" : "取消"}
            </button>
            <button
              type="button"
              className="primary-button"
              disabled={!canSave}
              onClick={() => void saveBinding()}
            >
              {saving
                ? isEn
                  ? "Registering…"
                  : "正在注册…"
                : conflictLabel
                ? isEn
                  ? "Replace and assign"
                  : "替换并绑定"
                : isEn
                  ? "Assign hotkey"
                  : "绑定快捷键"}
            </button>
          </div>
        </footer>
      </section>
    </div>
  );
}
