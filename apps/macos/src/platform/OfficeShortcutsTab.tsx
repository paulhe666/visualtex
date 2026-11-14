import { useState } from "react";
import { FileText, Pencil, Plus, Presentation, Trash2 } from "lucide-react";
import { OfficeHotkeyRecorderDialog } from "../components/OfficeHotkeyRecorderDialog";
import { formatFormulaHotkeyChord } from "../../../shared/visualtex-app/src/shortcuts/formulaHotkeys";
import { useOfficeHotkeyStore } from "../stores/officeHotkeyStore";
import { useEditorStore } from "../../../shared/visualtex-app/src/stores/editorStore";
import {
  OFFICE_HOTKEY_ACTIONS,
  officeHotkeyActionLabel,
  type OfficeHotkeyActionDefinition,
  type OfficeHotkeyHost,
} from "../shortcuts/officeHotkeys";

/**
 * Settings → Keyboard shortcuts → Office on macOS: system-wide hotkeys for the Word and PowerPoint add-ins' Ribbon
 * buttons, registered by the app.
 */
export function OfficeShortcutsTab() {
  const [editingOfficeAction, setEditingOfficeAction] =
    useState<OfficeHotkeyActionDefinition | null>(null);
  const officeBindings = useOfficeHotkeyStore((state) => state.bindings);
  const removeOfficeBinding = useOfficeHotkeyStore(
    (state) => state.removeBinding,
  );
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";

  const renderOfficeGroup = (
    host: OfficeHotkeyHost,
    title: string,
  ) => (
    <section className="office-hotkey-group" data-office-hotkey-host={host}>
      <header className="formula-hotkey-section-heading">
        {host === "word" ? <FileText size={16} /> : <Presentation size={16} />}
        <strong>{title}</strong>
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

  return (
    <>
      {renderOfficeGroup("word", "Microsoft Word")}
      {renderOfficeGroup("powerpoint", "Microsoft PowerPoint")}
      <OfficeHotkeyRecorderDialog
        action={editingOfficeAction}
        onClose={() => setEditingOfficeAction(null)}
      />
    </>
  );
}
