import { useEffect, useState } from "react";
import { Keyboard, X } from "lucide-react";
import {
  officeFormulaShortcutFromEvent,
  type OfficeFormulaShortcutCapture,
} from "../office/officeFormulaShortcuts";
import { useEditorStore } from "../stores/editorStore";
import type {
  OfficeFormulaHost,
  OfficeFormulaShortcutSettings,
} from "../types/formula";

type ShortcutKind = keyof OfficeFormulaShortcutSettings;

interface ShortcutTarget {
  host: OfficeFormulaHost;
  kind: ShortcutKind;
}

interface OfficeFormulaShortcutSettingsSectionProps {
  onOpenFormulaHotkeys: () => void;
}

function captureErrorText(
  result: Extract<OfficeFormulaShortcutCapture, { error: string }>,
  isEn: boolean,
) {
  switch (result.error) {
    case "modifier":
      return isEn
        ? "Include Ctrl or Alt in the shortcut. Shift can be added."
        : "快捷键必须包含 Ctrl 或 Alt，Shift 可以作为附加键。";
    case "reserved":
      return isEn
        ? "Ctrl+Alt+O is reserved for VisualTeX Quick OCR."
        : "Ctrl+Alt+O 已保留给 VisualTeX 快速 OCR。";
    default:
      return isEn
        ? "Use a letter, number, or F1–F24 together with Ctrl or Alt."
        : "请使用字母、数字或 F1–F24，并搭配 Ctrl 或 Alt。";
  }
}

function sameTarget(left: ShortcutTarget | null, right: ShortcutTarget) {
  return left?.host === right.host && left.kind === right.kind;
}

export function OfficeFormulaShortcutSettingsSection({
  onOpenFormulaHotkeys,
}: OfficeFormulaShortcutSettingsSectionProps) {
  const language = useEditorStore((state) => state.language);
  const wordSettings = useEditorStore((state) => state.wordFormulaShortcuts);
  const powerPointSettings = useEditorStore(
    (state) => state.powerPointFormulaShortcuts,
  );
  const setShortcut = useEditorStore((state) => state.setOfficeFormulaShortcut);
  const [capturing, setCapturing] = useState<ShortcutTarget | null>(null);
  const [message, setMessage] = useState("");
  const isEn = language === "en";

  useEffect(() => {
    if (!capturing) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (event.repeat || event.isComposing || event.key === "Process") return;
      if (
        event.key === "Escape" &&
        !event.ctrlKey &&
        !event.altKey &&
        !event.shiftKey &&
        !event.metaKey
      ) {
        setCapturing(null);
        setMessage("");
        return;
      }
      if (event.getModifierState("AltGraph")) {
        setMessage(
          isEn
            ? "AltGr is used for text input and cannot be assigned."
            : "AltGr 用于文字输入，不能设置为快捷键。",
        );
        return;
      }
      const result = officeFormulaShortcutFromEvent(event);
      if ("error" in result) {
        setMessage(captureErrorText(result, isEn));
        return;
      }
      const hostSettings =
        capturing.host === "word" ? wordSettings : powerPointSettings;
      const otherKind = capturing.kind === "inline" ? "display" : "inline";
      const moved = hostSettings[otherKind].shortcut === result.shortcut;
      setShortcut(capturing.host, capturing.kind, {
        shortcut: result.shortcut,
      });
      setCapturing(null);
      setMessage(
        moved
          ? isEn
            ? `${result.shortcut} was moved from the other action in this app.`
            : `${result.shortcut} 已从此应用的另一个动作移到这里。`
          : "",
      );
    };
    document.addEventListener("keydown", handleKeyDown, true);
    return () => document.removeEventListener("keydown", handleKeyDown, true);
  }, [capturing, isEn, powerPointSettings, setShortcut, wordSettings]);

  const rows: ReadonlyArray<{
    kind: ShortcutKind;
    labelEn: string;
    labelZh: string;
  }> = [
    { kind: "inline", labelEn: "Insert inline formula", labelZh: "插入行内公式" },
    { kind: "display", labelEn: "Insert display formula", labelZh: "插入行间公式" },
  ];
  const hosts: ReadonlyArray<{
    host: OfficeFormulaHost;
    label: string;
    descriptionEn: string;
    descriptionZh: string;
    settings: OfficeFormulaShortcutSettings;
  }> = [
    {
      host: "word",
      label: "Word",
      descriptionEn: "Calls the matching VisualTeX Ribbon action in Word.",
      descriptionZh: "调用 Word 中对应的 VisualTeX Ribbon 动作。",
      settings: wordSettings,
    },
    {
      host: "powerPoint",
      label: "PowerPoint",
      descriptionEn:
        "OMML follows the text caret; OLE creates an editable slide object.",
      descriptionZh: "OMML 跟随文本光标插入；OLE 创建可编辑的幻灯片对象。",
      settings: powerPointSettings,
    },
  ];

  return (
    <div className="settings-section" data-office-formula-shortcuts>
      <div className="settings-section-title">
        <Keyboard size={18} />
        <div>
          <h3>{isEn ? "Keyboard shortcuts" : "快捷键"}</h3>
          <p>
            {isEn
              ? "Configure Office insertion and editor formula shortcuts in one place."
              : "在同一处配置 Office 插入快捷键和编辑器公式快捷键。"}
          </p>
        </div>
      </div>

      <div className="office-formula-shortcut-hosts">
        {hosts.map(({ host, label, descriptionEn, descriptionZh, settings }) => (
          <section
            className="office-formula-shortcut-host"
            data-office-formula-shortcut-host={host}
            key={host}
          >
            <header className="office-formula-shortcut-host-header">
              <strong>{label}</strong>
              <small>{isEn ? descriptionEn : descriptionZh}</small>
            </header>
            <div className="office-formula-shortcut-list">
              {rows.map(({ kind, labelEn, labelZh }) => {
                const setting = settings[kind];
                const target = { host, kind } satisfies ShortcutTarget;
                const isCapturing = sameTarget(capturing, target);
                return (
                  <div
                    className="office-formula-shortcut-row"
                    data-office-formula-shortcut={kind}
                    key={kind}
                  >
                    <div className="office-formula-shortcut-copy">
                      <strong>{isEn ? labelEn : labelZh}</strong>
                      <small>
                        {isEn
                          ? `Overrides the same shortcut only in ${label}.`
                          : `仅覆盖 ${label} 中相同的快捷键。`}
                      </small>
                    </div>
                    <button
                      type="button"
                      className={`office-formula-shortcut-capture${isCapturing ? " is-capturing" : ""}`}
                      aria-pressed={isCapturing}
                      onClick={() => {
                        setMessage("");
                        setCapturing(isCapturing ? null : target);
                      }}
                    >
                      <kbd>
                        {isCapturing
                          ? isEn
                            ? "Press keys…"
                            : "请按键…"
                          : setting.shortcut ?? (isEn ? "None" : "无")}
                      </kbd>
                    </button>
                    <button
                      type="button"
                      className="office-formula-shortcut-clear"
                      disabled={!setting.shortcut && !isCapturing}
                      onClick={() => {
                        setShortcut(host, kind, { shortcut: null });
                        if (isCapturing) setCapturing(null);
                        setMessage("");
                      }}
                      aria-label={
                        isEn
                          ? `Clear ${label} ${labelEn}`
                          : `清除 ${label} ${labelZh}快捷键`
                      }
                      title={isEn ? "Set to None" : "设为无"}
                    >
                      <X size={14} />
                      {isEn ? "None" : "无"}
                    </button>
                    <label className="office-formula-format-select">
                      <span>{isEn ? "Format" : "格式"}</span>
                      <select
                        value={setting.format}
                        onChange={(event) =>
                          setShortcut(host, kind, {
                            format:
                              event.target.value === "omml" ? "omml" : "ole",
                          })
                        }
                        aria-label={
                          isEn
                            ? `${label} ${labelEn} format`
                            : `${label} ${labelZh}格式`
                        }
                      >
                        <option value="ole">OLE</option>
                        <option value="omml">OMML</option>
                      </select>
                    </label>
                  </div>
                );
              })}
            </div>
          </section>
        ))}
      </div>

      <p className="office-formula-shortcut-help">
        {message ||
          (capturing
            ? isEn
              ? "Press the shortcut now, or Esc to cancel."
              : "现在按下快捷键，或按 Esc 取消。"
            : isEn
              ? "Actions in one app cannot share a shortcut. Word and PowerPoint may reuse the same shortcut because they run in separate hosts."
              : "同一应用内的两个动作不能重复；Word 与 PowerPoint 属于不同宿主，可以使用相同组合键。")}
      </p>

      <div className="office-formula-editor-hotkeys">
        <div>
          <strong>{isEn ? "Editor formula hotkeys" : "编辑器公式快捷键"}</strong>
          <small>
            {isEn
              ? "Manage shortcuts that insert saved formulas inside the VisualTeX editor."
              : "管理仅在 VisualTeX 编辑器内插入常用公式的快捷键。"}
          </small>
        </div>
        <button
          type="button"
          className="secondary-button settings-hotkey-button"
          onClick={onOpenFormulaHotkeys}
        >
          <Keyboard size={15} />
          {isEn ? "Manage formula hotkeys" : "管理公式快捷键"}
        </button>
      </div>
    </div>
  );
}
