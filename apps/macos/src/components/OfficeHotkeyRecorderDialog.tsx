import { useEffect, useMemo, useRef, useState } from "react";
import { isTauri } from "@tauri-apps/api/core";
import {
  AlertTriangle,
  Check,
  FileText,
  Keyboard,
  Presentation,
  X,
} from "lucide-react";
import {
  formatFormulaHotkeyChord,
  formulaHotkeyChordFromEvent,
  formulaHotkeyChordId,
  formulaHotkeyHasModifier,
  formulaHotkeyTargetLabel,
  protectedFormulaHotkeyAction,
  type FormulaHotkeyChord,
} from "../shortcuts/formulaHotkeys";
import {
  officeHotkeyActionLabel,
  officeHotkeyAction,
  officeHotkeySupportsCode,
  type OfficeHotkeyActionDefinition,
} from "../shortcuts/officeHotkeys";
import { configureOfficeHotkeys } from "../runtime/officeHotkeys";
import { useEditorStore } from "../stores/editorStore";
import { useFormulaHotkeyStore } from "../stores/formulaHotkeyStore";
import { useOfficeHotkeyStore } from "../stores/officeHotkeyStore";

interface Props {
  action: OfficeHotkeyActionDefinition | null;
  onClose: () => void;
}

export function OfficeHotkeyRecorderDialog({ action, onClose }: Props) {
  const dialogRef = useRef<HTMLElement>(null);
  const language = useEditorStore((state) => state.language);
  const officeBindings = useOfficeHotkeyStore((state) => state.bindings);
  const setOfficeBinding = useOfficeHotkeyStore((state) => state.setBinding);
  const formulaBindings = useFormulaHotkeyStore((state) => state.bindings);
  const removeFormulaBinding = useFormulaHotkeyStore(
    (state) => state.removeBinding,
  );
  const [capturedChord, setCapturedChord] = useState<FormulaHotkeyChord | null>(
    null,
  );
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState("");
  const isEn = language === "en";
  const existingBinding = action
    ? officeBindings.find((binding) => binding.actionId === action.id) ?? null
    : null;

  useEffect(() => {
    setCapturedChord(existingBinding?.chord ?? null);
    setSaveError("");
  }, [action?.id, existingBinding?.updatedAt]);

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
      if (chord) {
        setCapturedChord(chord);
        setSaveError("");
      }
    };
    document.addEventListener("keydown", handleKeyDown, true);
    return () => {
      window.cancelAnimationFrame(frame);
      document.removeEventListener("keydown", handleKeyDown, true);
    };
  }, [action, onClose]);

  const conflicts = useMemo(() => {
    if (!action || !capturedChord) return { office: null, formula: null };
    const chordId = formulaHotkeyChordId(capturedChord);
    return {
      office:
        officeBindings.find(
          (binding) =>
            binding.actionId !== action.id &&
            formulaHotkeyChordId(binding.chord) === chordId,
        ) ?? null,
      formula:
        formulaBindings.find(
          (binding) => formulaHotkeyChordId(binding.chord) === chordId,
        ) ?? null,
    };
  }, [action, capturedChord, formulaBindings, officeBindings]);

  if (!action) return null;

  const protectedAction = capturedChord
    ? protectedFormulaHotkeyAction(capturedChord, language)
    : null;
  const hasModifier = capturedChord
    ? formulaHotkeyHasModifier(capturedChord)
    : false;
  const supportedCode = capturedChord
    ? officeHotkeySupportsCode(capturedChord.code)
    : false;
  const canSave = Boolean(
    capturedChord && hasModifier && supportedCode && !protectedAction && !saving,
  );
  const actionLabel = officeHotkeyActionLabel(action, language);
  const officeConflictLabel = conflicts.office
    ? officeHotkeyActionLabel(
        officeHotkeyAction(conflicts.office.actionId),
        language,
      )
    : null;

  const saveBinding = async () => {
    if (!capturedChord || !canSave) return;
    setSaving(true);
    setSaveError("");
    const chordId = formulaHotkeyChordId(capturedChord);
    const nextBindings = [
      { actionId: action.id, chord: capturedChord, updatedAt: Date.now() },
      ...officeBindings.filter(
        (binding) =>
          binding.actionId !== action.id &&
          formulaHotkeyChordId(binding.chord) !== chordId,
      ),
    ];
    try {
      if (isTauri()) await configureOfficeHotkeys(nextBindings);
      if (conflicts.formula) removeFormulaBinding(conflicts.formula.id);
      setOfficeBinding(action.id, capturedChord);
      onClose();
    } catch (error) {
      setSaveError(
        error instanceof Error
          ? error.message
          : String(
              error ||
                (isEn
                  ? "Unable to register this global hotkey."
                  : "无法注册该全局快捷键。"),
            ),
      );
    } finally {
      setSaving(false);
    }
  };

  const conflictLabel = officeConflictLabel ??
    (conflicts.formula
      ? formulaHotkeyTargetLabel(conflicts.formula.target, language)
      : null);

  return (
    <div
      className="modal-backdrop formula-hotkey-modal-backdrop"
      role="presentation"
      onMouseDown={onClose}
    >
      <section
        ref={dialogRef}
        className="formula-hotkey-recorder-dialog office-hotkey-recorder-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="office-hotkey-recorder-title"
        tabIndex={-1}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="dialog-header">
          <div>
            <span className="eyebrow">OFFICE HOT KEY</span>
            <h2 id="office-hotkey-recorder-title">
              {isEn ? "Set Office hotkey" : "设置 Office 快捷键"}
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
              {action.host === "word" ? (
                <FileText size={28} />
              ) : (
                <Presentation size={28} />
              )}
            </div>
            <div>
              <strong>{actionLabel}</strong>
              <span>{action.host === "word" ? "Microsoft Word" : "Microsoft PowerPoint"}</span>
              <code>
                {isEn ? "Ribbon button: " : "Ribbon 按钮："}
                {isEn ? action.ribbonLabelEn : action.ribbonLabelZh}
              </code>
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
                  ? "Requires Ctrl, Option/Alt or Command."
                  : "需要 Ctrl、Option/Alt 或 Command。"}
              </span>
            </div>
          )}
          {capturedChord && hasModifier && !supportedCode && (
            <div className="formula-hotkey-message is-danger" role="alert">
              <AlertTriangle size={16} />
              <span>
                {isEn
                  ? "This key cannot be registered as a macOS global shortcut."
                  : "该按键不能注册为 macOS 全局快捷键。"}
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
          {conflictLabel && !protectedAction && hasModifier && supportedCode && (
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
