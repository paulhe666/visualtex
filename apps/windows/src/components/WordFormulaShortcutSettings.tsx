import { useEffect, useState } from "react";
import { Keyboard, X } from "lucide-react";
import {
  wordFormulaShortcutFromEvent,
  type WordFormulaShortcutCapture,
} from "../office/wordFormulaShortcuts";
import { useEditorStore } from "../stores/editorStore";
import type { WordFormulaShortcutSettings } from "../types/formula";

type ShortcutKind = keyof WordFormulaShortcutSettings;

function captureErrorText(
  result: Extract<WordFormulaShortcutCapture, { error: string }>,
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

export function WordFormulaShortcutSettingsSection() {
  const language = useEditorStore((state) => state.language);
  const settings = useEditorStore((state) => state.wordFormulaShortcuts);
  const setShortcut = useEditorStore((state) => state.setWordFormulaShortcut);
  const [capturing, setCapturing] = useState<ShortcutKind | null>(null);
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
      const result = wordFormulaShortcutFromEvent(event);
      if ("error" in result) {
        setMessage(captureErrorText(result, isEn));
        return;
      }
      const otherKind = capturing === "inline" ? "display" : "inline";
      const moved = settings[otherKind].shortcut === result.shortcut;
      setShortcut(capturing, { shortcut: result.shortcut });
      setCapturing(null);
      setMessage(
        moved
          ? isEn
            ? `${result.shortcut} was moved from the other action.`
            : `${result.shortcut} 已从另一个动作移到这里。`
          : "",
      );
    };
    document.addEventListener("keydown", handleKeyDown, true);
    return () => document.removeEventListener("keydown", handleKeyDown, true);
  }, [capturing, isEn, setShortcut, settings]);

  const rows: ReadonlyArray<{
    kind: ShortcutKind;
    labelEn: string;
    labelZh: string;
  }> = [
    { kind: "inline", labelEn: "Insert inline formula", labelZh: "插入行内公式" },
    { kind: "display", labelEn: "Insert display formula", labelZh: "插入行间公式" },
  ];

  return (
    <div className="settings-section" data-word-formula-shortcuts>
      <div className="settings-section-title">
        <Keyboard size={18} />
        <div>
          <h3>{isEn ? "Word formula shortcuts" : "Word 公式快捷键"}</h3>
          <p>
            {isEn
              ? "Each shortcut calls the matching VisualTeX Ribbon action in Word."
              : "快捷键会直接调用 Word 中对应的 VisualTeX Ribbon 动作。"}
          </p>
        </div>
      </div>

      <div className="word-formula-shortcut-list">
        {rows.map(({ kind, labelEn, labelZh }) => {
          const setting = settings[kind];
          const isCapturing = capturing === kind;
          return (
            <div
              className="word-formula-shortcut-row"
              data-word-formula-shortcut={kind}
              key={kind}
            >
              <div className="word-formula-shortcut-copy">
                <strong>{isEn ? labelEn : labelZh}</strong>
                <small>
                  {isEn
                    ? "Overrides the same shortcut in Word while enabled."
                    : "启用后会覆盖 Word 中相同的快捷键。"}
                </small>
              </div>
              <button
                type="button"
                className={`word-formula-shortcut-capture${isCapturing ? " is-capturing" : ""}`}
                aria-pressed={isCapturing}
                onClick={() => {
                  setMessage("");
                  setCapturing(isCapturing ? null : kind);
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
                className="word-formula-shortcut-clear"
                disabled={!setting.shortcut && !isCapturing}
                onClick={() => {
                  setShortcut(kind, { shortcut: null });
                  if (capturing === kind) setCapturing(null);
                  setMessage("");
                }}
                aria-label={isEn ? `Clear ${labelEn}` : `清除${labelZh}快捷键`}
                title={isEn ? "Set to None" : "设为无"}
              >
                <X size={14} />
                {isEn ? "None" : "无"}
              </button>
              <label className="word-formula-format-select">
                <span>{isEn ? "Format" : "格式"}</span>
                <select
                  value={setting.format}
                  onChange={(event) =>
                    setShortcut(kind, {
                      format: event.target.value === "omml" ? "omml" : "ole",
                    })
                  }
                  aria-label={
                    isEn ? `${labelEn} format` : `${labelZh}格式`
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

      <p className="word-formula-shortcut-help">
        {message ||
          (capturing
            ? isEn
              ? "Press the shortcut now, or Esc to cancel."
              : "现在按下快捷键，或按 Esc 取消。"
            : isEn
              ? "Supported keys: letters, numbers, and F1–F24 with Ctrl or Alt."
              : "支持 Ctrl 或 Alt 搭配字母、数字及 F1–F24。")}
      </p>
    </div>
  );
}
