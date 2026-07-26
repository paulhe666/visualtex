import { useEffect, useMemo, useRef, useState } from "react";
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
import { useEditorStore } from "../stores/editorStore";

interface Props {
  target: FormulaHotkeyTarget | null;
  onClose: () => void;
}

export function FormulaHotkeyRecorderDialog({ target, onClose }: Props) {
  const dialogRef = useRef<HTMLElement>(null);
  const bindings = useFormulaHotkeyStore((state) => state.bindings);
  const setBinding = useFormulaHotkeyStore((state) => state.setBinding);
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const existingBinding = target
    ? bindings.find((binding) => binding.target.id === target.id) ?? null
    : null;
  const [capturedChord, setCapturedChord] = useState<FormulaHotkeyChord | null>(
    null,
  );

  useEffect(() => {
    setCapturedChord(existingBinding?.chord ?? null);
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

  const conflict = useMemo(() => {
    if (!target || !capturedChord) return null;
    const chordId = formulaHotkeyChordId(capturedChord);
    return (
      bindings.find(
        (binding) =>
          binding.target.id !== target.id &&
          formulaHotkeyChordId(binding.chord) === chordId,
      ) ?? null
    );
  }, [bindings, capturedChord, target]);

  if (!target) return null;

  const protectedAction = capturedChord
    ? protectedFormulaHotkeyAction(capturedChord, language)
    : null;
  const hasModifier = capturedChord
    ? formulaHotkeyHasModifier(capturedChord)
    : false;
  const canSave = Boolean(capturedChord && hasModifier && !protectedAction);
  const targetLabel = formulaHotkeyTargetLabel(target, language);

  const saveBinding = () => {
    if (!capturedChord || !canSave) return;
    setBinding(target, capturedChord);
    onClose();
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
                {isEn ? "Press a shortcut now" : "现在按下需要绑定的快捷键"}
              </strong>
            )}
            <span>
              {isEn
                ? "Include Ctrl, Option/Alt or Command. Press Esc to cancel."
                : "请加入 Ctrl、Option/Alt 或 Command；按 Esc 取消。"}
            </span>
          </div>

          {capturedChord && !hasModifier && (
            <div className="formula-hotkey-message is-warning" role="alert">
              <AlertTriangle size={16} />
              <span>
                {isEn
                  ? "This would interfere with normal formula input. Add Ctrl, Option/Alt or Command; Shift can be used in addition."
                  : "该组合会影响正常公式输入，请加入 Ctrl、Option/Alt 或 Command；Shift 可作为附加修饰键。"}
              </span>
            </div>
          )}

          {protectedAction && (
            <div className="formula-hotkey-message is-danger" role="alert">
              <AlertTriangle size={16} />
              <span>
                {isEn
                  ? `This shortcut is reserved for ${protectedAction} and cannot be overridden.`
                  : `该快捷键已保留用于“${protectedAction}”，不能被公式快捷键覆盖。`}
              </span>
            </div>
          )}

          {conflict && !protectedAction && hasModifier && (
            <div className="formula-hotkey-message is-warning" role="alert">
              <AlertTriangle size={16} />
              <span>
                {isEn
                  ? `Currently assigned to “${formulaHotkeyTargetLabel(conflict.target, language)}”. Saving will replace that assignment.`
                  : `当前已分配给“${formulaHotkeyTargetLabel(conflict.target, language)}”，保存后将替换原绑定。`}
              </span>
            </div>
          )}

          {capturedChord && canSave && !conflict && (
            <div className="formula-hotkey-message is-success">
              <Check size={16} />
              <span>
                {isEn
                  ? "Available in the visual formula editor."
                  : "该快捷键可在可视化公式编辑区使用。"}
              </span>
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
              {conflict
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
