import { useMemo, useState } from "react";
import { PanelLeftClose, Star } from "lucide-react";
import type { LatexCommand } from "../types/command";
import {
  categoryLabels,
  categoryLabelsEn,
  commandRegistry,
  commonCommandIds,
} from "../autocomplete/commandRegistry";
import { MathPreview } from "../components/MathPreview";
import { useEditorStore } from "../stores/editorStore";

interface Props {
  onInsert: (command: LatexCommand) => void;
  onClose: () => void;
}

const categories = [
  "common",
  "structure",
  "calculus",
  "matrix",
  "greek",
  "relation",
  "set",
  "arrow",
  "physics",
];

const previewSizeClass = (command: LatexCommand) => {
  const length = command.previewLatex.length;
  const wide =
    length > 22 ||
    command.previewLatex.includes("\\begin") ||
    command.id === "derivative" ||
    command.id === "partial";
  const compact = length > 42 || command.previewLatex.includes("cases");
  return (wide ? " is-wide" : "") + (compact ? " is-compact" : "");
};

export function FormulaToolbar({ onInsert, onClose }: Props) {
  const [activeCategory, setActiveCategory] = useState("common");
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";

  const visibleCommands = useMemo(() => {
    if (activeCategory === "common") {
      return commonCommandIds
        .map((id) => commandRegistry.find((command) => command.id === id))
        .filter((command): command is LatexCommand => Boolean(command));
    }
    return commandRegistry.filter(
      (command) => command.category === activeCategory,
    );
  }, [activeCategory]);

  return (
    <aside
      className="formula-toolbar"
      aria-label={isEn ? "Formula toolbar" : "公式工具栏"}
    >
      <header className="formula-toolbar-header">
        <strong>{isEn ? "Formula tools" : "公式工具"}</strong>
        <button
          type="button"
          className="icon-button compact"
          onClick={onClose}
          aria-label={isEn ? "Hide formula tools" : "隐藏公式工具"}
          title={isEn ? "Hide formula tools" : "隐藏公式工具"}
        >
          <PanelLeftClose size={16} />
        </button>
      </header>

      <nav className="toolbar-tabs" aria-label={isEn ? "Formula categories" : "公式分类"}>
        {categories.map((category) => (
          <button
            key={category}
            type="button"
            className={
              "toolbar-tab " +
              (activeCategory === category ? "is-active" : "")
            }
            aria-pressed={activeCategory === category}
            onClick={() => setActiveCategory(category)}
          >
            {category === "common" && <Star size={13} />}
            {(isEn ? categoryLabelsEn : categoryLabels)[category]}
          </button>
        ))}
      </nav>

      <div className="template-strip" aria-label={isEn ? "Formula templates" : "公式模板"}>
        {visibleCommands.map((command) => (
          <button
            type="button"
            className={"template-button" + previewSizeClass(command)}
            key={command.id}
            onClick={() => onInsert(command)}
            title={
              (isEn ? command.labelEn : command.labelZh) +
              " · " +
              command.command
            }
          >
            <MathPreview latex={command.previewLatex} />
            <span>{isEn ? command.labelEn : command.labelZh}</span>
          </button>
        ))}
      </div>

    </aside>
  );
}
