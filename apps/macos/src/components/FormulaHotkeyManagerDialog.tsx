import { useEffect, useMemo, useRef, useState } from "react";
import {
  FileText,
  Keyboard,
  Pencil,
  Plus,
  Presentation,
  Trash2,
  X,
} from "lucide-react";
import { MathPreview } from "./MathPreview";
import { FormulaHotkeyRecorderDialog } from "./FormulaHotkeyRecorderDialog";
import { OfficeHotkeyRecorderDialog } from "./OfficeHotkeyRecorderDialog";
import {
  formatFormulaHotkeyChord,
  formulaHotkeyTargetKindLabel,
  formulaHotkeyTargetLabel,
  type FormulaHotkeyTarget,
} from "../shortcuts/formulaHotkeys";
import { useFormulaHotkeyStore } from "../stores/formulaHotkeyStore";
import { useOfficeHotkeyStore } from "../stores/officeHotkeyStore";
import { useEditorStore } from "../stores/editorStore";
import {
  OFFICE_HOTKEY_ACTIONS,
  officeHotkeyActionLabel,
  type OfficeHotkeyActionDefinition,
  type OfficeHotkeyHost,
} from "../shortcuts/officeHotkeys";

interface Props {
  open: boolean;
  onClose: () => void;
}

export function FormulaHotkeyManagerDialog({ open, onClose }: Props) {
  const dialogRef = useRef<HTMLElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const [editingTarget, setEditingTarget] =
    useState<FormulaHotkeyTarget | null>(null);
  const [editingOfficeAction, setEditingOfficeAction] =
    useState<OfficeHotkeyActionDefinition | null>(null);
  const bindings = useFormulaHotkeyStore((state) => state.bindings);
  const removeBinding = useFormulaHotkeyStore((state) => state.removeBinding);
  const officeBindings = useOfficeHotkeyStore((state) => state.bindings);
  const removeOfficeBinding = useOfficeHotkeyStore(
    (state) => state.removeBinding,
  );
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const sortedBindings = useMemo(
    () => [...bindings].sort((left, right) => right.updatedAt - left.updatedAt),
    [bindings],
  );

  useEffect(() => {
    if (!open) {
      setEditingTarget(null);
      setEditingOfficeAction(null);
      return;
    }
    previousFocusRef.current = document.activeElement as HTMLElement | null;
    const frame = window.requestAnimationFrame(() => dialogRef.current?.focus());
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape" || editingTarget || editingOfficeAction) return;
      event.preventDefault();
      onClose();
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      window.cancelAnimationFrame(frame);
      document.removeEventListener("keydown", handleKeyDown);
      previousFocusRef.current?.focus({ preventScroll: true });
    };
  }, [open, editingOfficeAction, editingTarget, onClose]);

  const renderOfficeGroup = (
    host: OfficeHotkeyHost,
    title: string,
  ) => (
    <section className="office-hotkey-group" data-office-hotkey-host={host}>
      <header className="formula-hotkey-section-heading">
        {host === "word" ? <FileText size={16} /> : <Presentation size={16} />}
        <strong>{title}</strong>
        <span>
          {isEn ? "Matches the installed Ribbon buttons" : "与已安装的 Ribbon 按钮一致"}
        </span>
      </header>
      <div className="office-hotkey-action-list">
        {OFFICE_HOTKEY_ACTIONS.filter((action) => action.host === host).map(
          (action) => {
            const binding = officeBindings.find(
              (item) => item.actionId === action.id,
            );
            const label = officeHotkeyActionLabel(action, language);
            return (
              <article
                className="office-hotkey-action-row"
                key={action.id}
                data-office-hotkey-action={action.id}
              >
                <div className="office-hotkey-action-icon" aria-hidden="true">
                  {host === "word" ? (
                    <FileText size={18} />
                  ) : (
                    <Presentation size={18} />
                  )}
                </div>
                <div className="formula-hotkey-binding-copy">
                  <strong>{label}</strong>
                  <span>
                    {isEn ? "Ribbon button: " : "Ribbon 按钮："}
                    {isEn ? action.ribbonLabelEn : action.ribbonLabelZh}
                  </span>
                </div>
                {binding ? (
                  <kbd>{formatFormulaHotkeyChord(binding.chord)}</kbd>
                ) : (
                  <span className="office-hotkey-unassigned">
                    {isEn ? "Not set" : "未设置"}
                  </span>
                )}
                <div className="formula-hotkey-binding-actions">
                  <button
                    type="button"
                    className="icon-button compact"
                    onClick={() => setEditingOfficeAction(action)}
                    aria-label={
                      binding
                        ? isEn
                          ? `Change hotkey for ${label}`
                          : `修改${label}的快捷键`
                        : isEn
                          ? `Set hotkey for ${label}`
                          : `设置${label}的快捷键`
                    }
                    title={binding ? (isEn ? "Change hotkey" : "修改快捷键") : (isEn ? "Set hotkey" : "设置快捷键")}
                  >
                    {binding ? <Pencil size={14} /> : <Plus size={14} />}
                  </button>
                  {binding && (
                    <button
                      type="button"
                      className="icon-button compact is-danger"
                      onClick={() => removeOfficeBinding(action.id)}
                      aria-label={
                        isEn
                          ? `Remove hotkey for ${label}`
                          : `删除${label}的快捷键`
                      }
                      title={isEn ? "Remove hotkey" : "删除快捷键"}
                    >
                      <Trash2 size={14} />
                    </button>
                  )}
                </div>
              </article>
            );
          },
        )}
      </div>
    </section>
  );

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
                  {sortedBindings.length + officeBindings.length}{" "}
                  {isEn ? "assigned" : "项已设置"}
                </strong>
                <small>
                  {isEn
                    ? "Office actions are global; editor shortcuts remain scoped to the formula editor."
                    : "Office 动作为全局快捷键；公式输入快捷键仍仅作用于公式编辑区。"}
                </small>
              </span>
            </div>
          </div>

          <div className="formula-hotkey-manager-content">
            {renderOfficeGroup("word", "Microsoft Word")}
            {renderOfficeGroup("powerpoint", "Microsoft PowerPoint")}
            <section className="formula-editor-hotkey-group">
              <header className="formula-hotkey-section-heading">
                <Keyboard size={16} />
                <strong>{isEn ? "Formula editor" : "公式编辑器"}</strong>
                <span>
                  {isEn
                    ? "Right-click a formula tool or tile to add more"
                    : "右键公式工具或磁贴可继续添加"}
                </span>
              </header>
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
            </section>
          </div>

          <footer className="dialog-footer">
            <span>
              {isEn
                ? "Office hotkeys run only while the matching Office app is frontmost."
                : "Office 快捷键仅在对应 Office 应用位于前台时执行。"}
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
      <OfficeHotkeyRecorderDialog
        action={editingOfficeAction}
        onClose={() => setEditingOfficeAction(null)}
      />
    </>
  );
}
