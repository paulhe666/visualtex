import { useEffect, useMemo, useRef, useState } from "react";
import { Keyboard, Pencil, Trash2, X } from "lucide-react";
import { MathPreview } from "./MathPreview";
import { FormulaHotkeyRecorderDialog } from "./FormulaHotkeyRecorderDialog";
import {
  formatFormulaHotkeyChord,
  formulaHotkeyTargetKindLabel,
  formulaHotkeyTargetLabel,
  type FormulaHotkeyTarget,
} from "../shortcuts/formulaHotkeys";
import { useFormulaHotkeyStore } from "../stores/formulaHotkeyStore";
import { useEditorStore } from "../stores/editorStore";

interface Props {
  open: boolean;
  onClose: () => void;
}

export function FormulaHotkeyManagerDialog({ open, onClose }: Props) {
  const dialogRef = useRef<HTMLElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const [editingTarget, setEditingTarget] =
    useState<FormulaHotkeyTarget | null>(null);
  const bindings = useFormulaHotkeyStore((state) => state.bindings);
  const removeBinding = useFormulaHotkeyStore((state) => state.removeBinding);
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const sortedBindings = useMemo(
    () => [...bindings].sort((left, right) => right.updatedAt - left.updatedAt),
    [bindings],
  );

  useEffect(() => {
    if (!open) {
      setEditingTarget(null);
      return;
    }
    previousFocusRef.current = document.activeElement as HTMLElement | null;
    const frame = window.requestAnimationFrame(() => dialogRef.current?.focus());
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape" || editingTarget) return;
      event.preventDefault();
      onClose();
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      window.cancelAnimationFrame(frame);
      document.removeEventListener("keydown", handleKeyDown);
      previousFocusRef.current?.focus({ preventScroll: true });
    };
  }, [open, editingTarget, onClose]);

  if (!open) return null;

  return (
    <>
      <div className="modal-backdrop formula-hotkey-modal-backdrop" role="presentation" onMouseDown={onClose}>
        <section
          ref={dialogRef}
          className="formula-hotkey-manager-dialog"
          role="dialog"
          aria-modal="true"
          aria-labelledby="formula-hotkey-manager-title"
          tabIndex={-1}
          onMouseDown={(event) => event.stopPropagation()}
        >
          <header className="dialog-header">
            <div>
              <span className="eyebrow">HOT KEYS</span>
              <h2 id="formula-hotkey-manager-title">
                {isEn ? "Formula hotkeys" : "公式快捷键"}
              </h2>
            </div>
            <button
              type="button"
              className="icon-button"
              onClick={onClose}
              aria-label={isEn ? "Close hotkey settings" : "关闭快捷键设置"}
            >
              <X size={18} />
            </button>
          </header>

          <div className="formula-hotkey-manager-summary">
            <div>
              <Keyboard size={18} />
              <span>
                <strong>
                  {sortedBindings.length} {isEn ? "assigned" : "项已设置"}
                </strong>
                <small>
                  {isEn
                    ? "Right-click a formula tool or tile to add another hotkey."
                    : "右键公式工具或磁贴，可以继续添加快捷键。"}
                </small>
              </span>
            </div>
          </div>

          <div className="formula-hotkey-manager-content">
            {sortedBindings.length === 0 ? (
              <div className="formula-hotkey-empty-state">
                <Keyboard size={28} />
                <strong>{isEn ? "No formula hotkeys yet" : "还没有设置公式快捷键"}</strong>
                <span>
                  {isEn
                    ? "Close this window, then right-click any formula tool, common tile or custom tile."
                    : "关闭此窗口后，右键任意公式工具、常用磁贴或自定义磁贴即可设置。"}
                </span>
              </div>
            ) : (
              <div className="formula-hotkey-binding-list">
                {sortedBindings.map((binding) => (
                  <article className="formula-hotkey-binding-row" key={binding.id}>
                    <div className="formula-hotkey-binding-preview">
                      <MathPreview latex={binding.target.command.previewLatex} fit />
                    </div>
                    <div className="formula-hotkey-binding-copy">
                      <strong>
                        {formulaHotkeyTargetLabel(binding.target, language)}
                      </strong>
                      <span>
                        {formulaHotkeyTargetKindLabel(binding.target, language)}
                        {" · "}
                        <code>{binding.target.command.command}</code>
                      </span>
                    </div>
                    <kbd>{formatFormulaHotkeyChord(binding.chord)}</kbd>
                    <div className="formula-hotkey-binding-actions">
                      <button
                        type="button"
                        className="icon-button compact"
                        onClick={() => setEditingTarget(binding.target)}
                        aria-label={
                          isEn
                            ? `Change hotkey for ${formulaHotkeyTargetLabel(binding.target, language)}`
                            : `修改${formulaHotkeyTargetLabel(binding.target, language)}的快捷键`
                        }
                        title={isEn ? "Change hotkey" : "修改快捷键"}
                      >
                        <Pencil size={14} />
                      </button>
                      <button
                        type="button"
                        className="icon-button compact is-danger"
                        onClick={() => removeBinding(binding.id)}
                        aria-label={
                          isEn
                            ? `Remove hotkey for ${formulaHotkeyTargetLabel(binding.target, language)}`
                            : `删除${formulaHotkeyTargetLabel(binding.target, language)}的快捷键`
                        }
                        title={isEn ? "Remove hotkey" : "删除快捷键"}
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  </article>
                ))}
              </div>
            )}
          </div>

          <footer className="dialog-footer">
            <span>
              {isEn
                ? "Hotkeys only take priority inside the visual formula editor."
                : "快捷键仅在可视化公式编辑区优先生效。"}
            </span>
            <button type="button" className="primary-button" onClick={onClose}>
              {isEn ? "Done" : "完成"}
            </button>
          </footer>
        </section>
      </div>

      <FormulaHotkeyRecorderDialog
        target={editingTarget}
        onClose={() => setEditingTarget(null)}
      />
    </>
  );
}
