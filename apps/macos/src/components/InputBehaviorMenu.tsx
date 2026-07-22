import { useEffect, useRef, useState } from "react";
import { ChevronDown, MousePointerClick } from "lucide-react";
import type { InputBehaviorSettingKey } from "../types/formula";
import { useEditorStore } from "../stores/editorStore";

interface InputBehaviorOption {
  key: InputBehaviorSettingKey;
  titleZh: string;
  titleEn: string;
  descriptionZh: string;
  descriptionEn: string;
}

const INPUT_BEHAVIOR_OPTIONS: InputBehaviorOption[] = [
  {
    key: "autoExitSuperscript",
    titleZh: "上标输入后跳出",
    titleEn: "Exit superscript after input",
    descriptionZh: "输入一个字符或一个工具栏符号后返回主公式区域",
    descriptionEn: "Return to the main formula after one character or toolbar symbol",
  },
  {
    key: "autoExitSubscript",
    titleZh: "下标输入后跳出",
    titleEn: "Exit subscript after input",
    descriptionZh: "输入一个字符或一个工具栏符号后返回主公式区域",
    descriptionEn: "Return to the main formula after one character or toolbar symbol",
  },
  {
    key: "autoExitAccent",
    titleZh: "重音内容输入后跳出",
    titleEn: "Exit accent after input",
    descriptionZh: "适用于 hat、bar、vec、tilde、dot 等包裹结构",
    descriptionEn: "Applies to hat, bar, vec, tilde, dot and similar accents",
  },
];

export function InputBehaviorMenu() {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);
  const language = useEditorStore((state) => state.language);
  const inputBehavior = useEditorStore((state) => state.inputBehavior);
  const setInputBehavior = useEditorStore((state) => state.setInputBehavior);
  const isEn = language === "en";

  useEffect(() => {
    if (!open) return;
    const handlePointerDown = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setOpen(false);
    };
    document.addEventListener("pointerdown", handlePointerDown, true);
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("pointerdown", handlePointerDown, true);
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [open]);

  return (
    <div ref={rootRef} className="input-behavior-menu">
      <button
        type="button"
        className={`canvas-input-behavior-trigger${open ? " is-active" : ""}`}
        aria-haspopup="dialog"
        aria-expanded={open}
        onClick={() => setOpen((value) => !value)}
        title={isEn ? "Input behavior" : "操作逻辑"}
      >
        <MousePointerClick size={14} />
        <span>{isEn ? "Input behavior" : "操作逻辑"}</span>
        <ChevronDown size={13} aria-hidden="true" />
      </button>

      {open && (
        <div
          className="input-behavior-popover"
          role="dialog"
          aria-label={isEn ? "Input behavior settings" : "操作逻辑设置"}
        >
          <div className="input-behavior-heading">
            <strong>{isEn ? "Caret auto-exit" : "光标自动跳出"}</strong>
            <span>
              {isEn
                ? "Choose which one-slot structures return to the main formula after input."
                : "分别选择哪些单槽结构在输入完成后自动返回主公式区域。"}
            </span>
          </div>

          <div className="input-behavior-options">
            {INPUT_BEHAVIOR_OPTIONS.map((option) => (
              <label className="input-behavior-option" key={option.key}>
                <span>
                  <strong>{isEn ? option.titleEn : option.titleZh}</strong>
                  <small>
                    {isEn ? option.descriptionEn : option.descriptionZh}
                  </small>
                </span>
                <input
                  type="checkbox"
                  checked={inputBehavior[option.key]}
                  onChange={(event) =>
                    setInputBehavior(option.key, event.target.checked)
                  }
                />
                <span className="input-behavior-switch" aria-hidden="true" />
              </label>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
