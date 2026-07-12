import { memo, useEffect, useRef } from "react";
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

interface SuggestionItemProps {
  command: LatexCommand;
  index: number;
  selected: boolean;
  pinned: boolean;
  isEn: boolean;
  selectedItemRef: React.RefObject<HTMLButtonElement | null>;
  onHighlight: (index: number) => void;
  onCommit: (command: LatexCommand) => void;
}

const SuggestionItem = memo(function SuggestionItem({
  command,
  index,
  selected,
  pinned,
  isEn,
  selectedItemRef,
  onHighlight,
  onCommit,
}: SuggestionItemProps) {
  return (
    <button
      ref={selected ? selectedItemRef : undefined}
      type="button"
      role="option"
      aria-selected={selected}
      className={"suggestion-item " + (selected ? "is-selected" : "")}
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
      <span className="suggestion-pin-slot" aria-hidden={!pinned}>
        {pinned ? <Pin size={13} aria-label={isEn ? "Pinned" : "已固定"} /> : null}
      </span>
      <CornerDownLeft
        size={15}
        className={"suggestion-enter " + (selected ? "is-visible" : "")}
        aria-hidden="true"
      />
    </button>
  );
});

export function CommandSuggestionPopup({
  suggestions,
  selectedIndex,
  position,
  onHighlight,
  onCommit,
  usage,
}: Props) {
  const language = useEditorStore((state) => state.language);
  const selectedItemRef = useRef<HTMLButtonElement>(null);
  const isEn = language === "en";

  useEffect(() => {
    selectedItemRef.current?.scrollIntoView({ block: "nearest" });
  }, [selectedIndex]);

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
          <SuggestionItem
            key={command.id}
            command={command}
            index={index}
            selected={index === selectedIndex}
            pinned={Boolean(usage[command.id]?.pinned)}
            isEn={isEn}
            selectedItemRef={selectedItemRef}
            onHighlight={onHighlight}
            onCommit={onCommit}
          />
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
