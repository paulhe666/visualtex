import { useEffect, useState } from "react";
import { Keyboard, X } from "lucide-react";
import {
  officeFormulaShortcutFromEvent,
  type OfficeFormulaShortcutCapture,
} from "../office/officeFormulaShortcuts";
import { useEditorStore } from "../stores/editorStore";
import type {
  PowerPointFormulaShortcutAction,
  WordFormulaShortcutSettings,
} from "../types/formula";

type WordShortcutAction = keyof WordFormulaShortcutSettings;
type ShortcutTarget =
  | { host: "word"; action: WordShortcutAction }
  | { host: "powerPoint"; action: PowerPointFormulaShortcutAction };

const powerPointActions: ReadonlyArray<{
  action: PowerPointFormulaShortcutAction;
  labelEn: string;
  labelZh: string;
}> = [
  {
    action: "svg",
    labelEn: "New SVG formula",
    labelZh: "新建 SVG 公式",
  },
  {
    action: "ole",
    labelEn: "New OLE formula",
    labelZh: "新建 OLE 公式",
  },
  {
    action: "inlineOmml",
    labelEn: "Insert inline OMML formula",
    labelZh: "插入行内 OMML 公式",
  },
  {
    action: "displayOmml",
    labelEn: "Insert display OMML formula",
    labelZh: "插入行间 OMML 公式",
  },
];

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
  return left?.host === right.host && left.action === right.action;
}

export function OfficeFormulaShortcutSettingsSection() {
  const language = useEditorStore((state) => state.language);
  const wordSettings = useEditorStore((state) => state.wordFormulaShortcuts);
  const powerPointSettings = useEditorStore(
    (state) => state.powerPointFormulaShortcuts,
  );
  const setWordShortcut = useEditorStore(
    (state) => state.setWordFormulaShortcut,
  );
  const setPowerPointShortcut = useEditorStore(
    (state) => state.setPowerPointFormulaShortcut,
  );
  const [capturing, setCapturing] = useState<ShortcutTarget | null>(null);
  const [message, setMessage] = useState("");
  const isEn = language === "en";

  const assignShortcut = (target: ShortcutTarget, shortcut: string | null) => {
    if (target.host === "word") {
      setWordShortcut(target.action, { shortcut });
    } else {
      setPowerPointShortcut(target.action, shortcut);
    }
  };

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
      if (capturing.host === "word" && !event.ctrlKey) {
        setMessage(isEn
          ? "Word shortcuts must include Ctrl; Alt-only shortcuts remain available to Word menus."
          : "Word 公式快捷键必须包含 Ctrl；仅含 Alt 的组合键保留给 Word 菜单。");
        return;
      }
      if (capturing.host === "word" && /^F(?:1[7-9]|2[0-4])$/.test(event.code)) {
        setMessage(isEn ? "Word supports F1–F16 here." : "这里的 Word 快捷键支持 F1–F16。");
        return;
      }
      const moved =
        capturing.host === "word"
          ? Object.entries(wordSettings).some(
              ([action, setting]) =>
                action !== capturing.action && setting.shortcut === result.shortcut,
            )
          : Object.entries(powerPointSettings).some(
              ([action, shortcut]) =>
                action !== capturing.action && shortcut === result.shortcut,
            );
      assignShortcut(capturing, result.shortcut);
      setCapturing(null);
      setMessage(
        moved
          ? isEn
            ? `${result.shortcut} was moved from another action in this app.`
            : `${result.shortcut} 已从此应用的另一个动作移到这里。`
          : "",
      );
    };
    document.addEventListener("keydown", handleKeyDown, true);
    return () => document.removeEventListener("keydown", handleKeyDown, true);
  }, [capturing, isEn, powerPointSettings, wordSettings]);

  const shortcutControls = (
    target: ShortcutTarget,
    label: string,
    shortcut: string | null,
  ) => {
    const isCapturing = sameTarget(capturing, target);
    return (
      <>
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
              : shortcut ?? (isEn ? "None" : "无")}
          </kbd>
        </button>
        <button
          type="button"
          className="office-formula-shortcut-clear"
          disabled={!shortcut && !isCapturing}
          onClick={() => {
            assignShortcut(target, null);
            if (isCapturing) setCapturing(null);
            setMessage("");
          }}
          aria-label={isEn ? `Clear ${label}` : `清除${label}快捷键`}
          title={isEn ? "Set to None" : "设为无"}
        >
          <X size={14} />
          {isEn ? "None" : "无"}
        </button>
      </>
    );
  };

  const wordRows: ReadonlyArray<{
    action: WordShortcutAction;
    labelEn: string;
    labelZh: string;
  }> = [
    { action: "inline", labelEn: "Insert inline formula", labelZh: "插入行内公式" },
    { action: "display", labelEn: "Insert display formula", labelZh: "插入行间公式" },
  ];

  return (
    <section className="shortcut-settings-group" data-office-formula-shortcuts>
      <div className="settings-section-title">
        <Keyboard size={18} />
        <div>
          <h3>{isEn ? "Office shortcuts" : "Office 快捷键"}</h3>
        </div>
      </div>

      <div className="office-formula-shortcut-hosts">
        <section
          className="office-formula-shortcut-host"
          data-office-formula-shortcut-host="word"
        >
          <header className="office-formula-shortcut-host-header">
            <strong>Word</strong>
          </header>
          <p className="office-formula-shortcut-help">
            {isEn
              ? "Word’s built-in and custom shortcuts take priority. Use an unassigned Ctrl combination in the document; clear a shortcut to disable it."
              : "Word 自带和自定义快捷键优先。请在正文中使用未被占用的 Ctrl 组合键；设为“无”即可停用。"}
          </p>
          <div className="office-formula-shortcut-list">
            {wordRows.map(({ action, labelEn, labelZh }) => {
              const setting = wordSettings[action];
              const target = { host: "word", action } as const;
              const label = isEn ? labelEn : labelZh;
              return (
                <div
                  className="office-formula-shortcut-row"
                  data-office-formula-shortcut={action}
                  key={action}
                >
                  <div className="office-formula-shortcut-copy">
                    <strong>{label}</strong>
                  </div>
                  {shortcutControls(target, `Word ${label}`, setting.shortcut)}
                  <label className="office-formula-format-select">
                    <span>{isEn ? "Format" : "格式"}</span>
                    <select
                      value={setting.format}
                      onChange={(event) =>
                        setWordShortcut(action, {
                          format: event.target.value === "omml" ? "omml" : "ole",
                        })
                      }
                      aria-label={isEn ? `Word ${label} format` : `Word ${label}格式`}
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

        <section
          className="office-formula-shortcut-host"
          data-office-formula-shortcut-host="powerPoint"
        >
          <header className="office-formula-shortcut-host-header">
            <strong>PowerPoint</strong>
          </header>
          <div className="office-formula-shortcut-list">
            {powerPointActions.map(
              ({ action, labelEn, labelZh }) => {
                const target = { host: "powerPoint", action } as const;
                const label = isEn ? labelEn : labelZh;
                return (
                  <div
                    className="office-formula-shortcut-row is-format-fixed"
                    data-office-formula-shortcut={action}
                    key={action}
                  >
                    <div className="office-formula-shortcut-copy">
                      <strong>{label}</strong>
                    </div>
                    {shortcutControls(target, `PowerPoint ${label}`, powerPointSettings[action])}
                  </div>
                );
              },
            )}
          </div>
        </section>
      </div>

      {message && <p className="office-formula-shortcut-help" role="status">{message}</p>}
    </section>
  );
}
