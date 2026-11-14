import { useMemo, useState } from "react";
import { Keyboard, Pencil, RotateCcw, Rows3, Trash2 } from "lucide-react";
import { MathPreview } from "./MathPreview";
import { FormulaHotkeyRecorderDialog } from "./FormulaHotkeyRecorderDialog";
import { FormulaRowHotkeyRecorderDialog } from "./FormulaRowHotkeyRecorderDialog";
import { ButtonShortcutSettingsSection } from "./ButtonShortcutSettings";
import { AnywhereHotkeySettingsSection } from "./AnywhereHotkeySettings";
import { ConfirmEscapeKeySettingsSection } from "./ConfirmEscapeKeySettings";
import {
  formatFormulaHotkeyChord,
  formulaHotkeyTargetKindLabel,
  formulaHotkeyTargetLabel,
  type FormulaHotkeyTarget,
} from "../shortcuts/formulaHotkeys";
import { useFormulaHotkeyStore } from "../stores/formulaHotkeyStore";
import { useFormulaRowHotkeyStore } from "../stores/formulaRowHotkeyStore";
import { useEditorStore } from "../stores/editorStore";
// each system's Office shortcuts: system-wide hotkeys on macOS, the add-ins' own shortcuts on Windows
import { OfficeShortcutsTab } from "@platform/OfficeShortcutsTab";
import {
  FORMULA_ROW_HOTKEY_ACTIONS,
  formulaRowHotkeyActionLabel,
  isDefaultFormulaRowHotkeyChord,
  type FormulaRowHotkeyActionDefinition,
} from "../shortcuts/formulaRowHotkeys";

const SHORTCUT_TABS = [
  { id: "buttons", en: "Buttons", zh: "按钮" },
  { id: "global", en: "Formula input", zh: "公式输入法" },
  { id: "office", en: "Office", zh: "Office" },
  { id: "editor", en: "Formula editor", zh: "公式编辑" },
] as const;

/** Settings → Keyboard shortcuts: every shortcut group, with the recorders they open. */
export function KeyboardShortcutsPanel() {
  const [tab, setTab] = useState<(typeof SHORTCUT_TABS)[number]["id"]>("buttons");
  const [editingTarget, setEditingTarget] =
    useState<FormulaHotkeyTarget | null>(null);
  const [editingRowAction, setEditingRowAction] =
    useState<FormulaRowHotkeyActionDefinition | null>(null);
  const bindings = useFormulaHotkeyStore((state) => state.bindings);
  const removeBinding = useFormulaHotkeyStore((state) => state.removeBinding);
  const rowBindings = useFormulaRowHotkeyStore((state) => state.bindings);
  const removeRowBinding = useFormulaRowHotkeyStore(
    (state) => state.removeBinding,
  );
  const resetRowBinding = useFormulaRowHotkeyStore(
    (state) => state.resetBinding,
  );
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const sortedBindings = useMemo(
    () => [...bindings].sort((left, right) => right.updatedAt - left.updatedAt),
    [bindings],
  );

  const renderRowGroup = () => (
    <section className="office-hotkey-group formula-row-hotkey-group">
      <header className="formula-hotkey-section-heading">
        <Rows3 size={16} />
        <strong>{isEn ? "Formula rows" : "公式行"}</strong>
      </header>
      <div className="office-hotkey-action-list">
        {FORMULA_ROW_HOTKEY_ACTIONS.map((action) => {
          const chord =
            rowBindings.find((item) => item.actionId === action.id)?.chord ??
            null;
          const label = formulaRowHotkeyActionLabel(action, language);
          const isDefault = isDefaultFormulaRowHotkeyChord(action.id, chord);
          return (
            <article
              className="office-hotkey-action-row"
              key={action.id}
              data-formula-row-hotkey-action={action.id}
            >
              <div className="office-hotkey-action-icon" aria-hidden="true">
                <Rows3 size={18} />
              </div>
              <div className="formula-hotkey-binding-copy">
                <strong>{label}</strong>
              </div>
              {chord ? (
                <kbd>{formatFormulaHotkeyChord(chord)}</kbd>
              ) : (
                <span className="office-hotkey-unassigned">
                  {isEn ? "Off" : "已关闭"}
                </span>
              )}
              <div className="formula-hotkey-binding-actions">
                <button
                  type="button"
                  className="icon-button compact"
                  onClick={() => setEditingRowAction(action)}
                  aria-label={
                    isEn
                      ? `Change hotkey for ${label}`
                      : `修改${label}的快捷键`
                  }
                  title={isEn ? "Change hotkey" : "修改快捷键"}
                >
                  <Pencil size={14} />
                </button>
                {!isDefault && (
                  <button
                    type="button"
                    className="icon-button compact"
                    onClick={() => resetRowBinding(action.id)}
                    aria-label={
                      isEn
                        ? `Restore the default hotkey for ${label}`
                        : `恢复${label}的默认快捷键`
                    }
                    title={isEn ? "Restore default" : "恢复默认"}
                  >
                    <RotateCcw size={14} />
                  </button>
                )}
                {chord && (
                  <button
                    type="button"
                    className="icon-button compact is-danger"
                    onClick={() => removeRowBinding(action.id)}
                    aria-label={
                      isEn
                        ? `Turn off the hotkey for ${label}`
                        : `关闭${label}的快捷键`
                    }
                    title={isEn ? "Turn off" : "关闭快捷键"}
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

  return (
    <>
      <div className="shortcut-manager-scroll">
        <div className="theme-segment shortcut-tabs" role="tablist" aria-label={isEn ? "Shortcut groups" : "快捷键分组"}>
          {SHORTCUT_TABS.map((item) => (
            <button
              key={item.id}
              type="button"
              role="tab"
              className={tab === item.id ? "is-active" : ""}
              aria-selected={tab === item.id}
              data-shortcut-tab={item.id}
              onClick={() => setTab(item.id)}
            >
              {isEn ? item.en : item.zh}
            </button>
          ))}
        </div>
        {tab === "buttons" && <ButtonShortcutSettingsSection />}
        {tab === "global" && <AnywhereHotkeySettingsSection />}
        {tab === "office" && <OfficeShortcutsTab />}
        {tab === "editor" && (
          <>
            {renderRowGroup()}
            <ConfirmEscapeKeySettingsSection />
            <section className="formula-editor-hotkey-group">
              <header className="formula-hotkey-section-heading">
                <Keyboard size={16} />
                <strong>{isEn ? "Formula editor" : "公式编辑器"}</strong>
              </header>
            {sortedBindings.length === 0 ? (
              <div className="formula-hotkey-empty-state">
                <Keyboard size={28} />
                <strong>{isEn ? "No formula hotkeys" : "未设置公式快捷键"}</strong>
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
          </>
        )}
      </div>

      <FormulaHotkeyRecorderDialog
        target={editingTarget}
        onClose={() => setEditingTarget(null)}
      />
      <FormulaRowHotkeyRecorderDialog
        action={editingRowAction}
        onClose={() => setEditingRowAction(null)}
      />
    </>
  );
}
