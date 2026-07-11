import { CornerDownLeft, Pin } from "lucide-react";
import type { LatexCommand } from "../types/command";
import { MathPreview } from "../components/MathPreview";
import { useEditorStore } from "../stores/editorStore";

interface Props {
  suggestions: LatexCommand[];
  selectedIndex: number;
  position: { left: number; top: number };
  onHighlight: (index: number) => void;
  onCommit: (command: LatexCommand) => void;
  usage: Record<string, { pinned: boolean; useCount: number }>;
}

export function CommandSuggestionPopup({
  suggestions,
  selectedIndex,
  position,
  onHighlight,
  onCommit,
  usage,
}: Props) {
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";

  if (!suggestions.length) return null;

  return (
    <div
      className="suggestion-popup"
      style={{ left: position.left, top: position.top }}
      role="listbox"
      aria-label={isEn ? "LaTeX command suggestions" : "LaTeX 命令候选"}
    >
      <div className="suggestion-header">
        <span>{isEn ? "Command suggestions" : "命令候选"}</span>
        <span className="suggestion-key-hint">
          {isEn ? "↑↓ Select · ↵ Insert" : "↑↓ 选择 · ↵ 插入"}
        </span>
      </div>
      <div className="suggestion-list">
        {suggestions.map((command, index) => (
          <button
            type="button"
            key={command.id}
            role="option"
            aria-selected={index === selectedIndex}
            className={"suggestion-item " + (index === selectedIndex ? "is-selected" : "")}
            onMouseEnter={() => onHighlight(index)}
            onMouseDown={(event) => {
              event.preventDefault();
              onHighlight(index);
            }}
            onDoubleClick={() => onCommit(command)}
          >
            <span className="suggestion-preview">
              <MathPreview latex={command.previewLatex} />
            </span>
            <span className="suggestion-copy">
              <span className="suggestion-command">{command.command}</span>
              <span className="suggestion-label">
                {isEn ? command.labelEn : command.labelZh}
              </span>
            </span>
            {usage[command.id]?.pinned && (
              <Pin size={13} aria-label={isEn ? "Pinned" : "已固定"} />
            )}
            {index === selectedIndex && <CornerDownLeft size={15} className="suggestion-enter" />}
          </button>
        ))}
      </div>
      <div className="suggestion-footer">
        {isEn
          ? "Click to select · Double-click to insert"
          : "单击选择 · 双击插入"}
      </div>
    </div>
  );
}
