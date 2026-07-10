import { useMemo, useState } from "react";
import { ChevronDown, ChevronUp, Star } from "lucide-react";
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

export function FormulaToolbar({ onInsert }: Props) {
  const [activeCategory, setActiveCategory] = useState("common");
  const [expanded, setExpanded] = useState(true);
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
    <section
      className={
        "formula-toolbar " + (expanded ? "is-expanded" : "is-collapsed")
      }
      aria-label={isEn ? "Formula toolbar" : "公式工具栏"}
    >
      <div className="toolbar-tabs">
        {categories.map((category) => (
          <button
            key={category}
            type="button"
            className={
              "toolbar-tab " +
              (activeCategory === category ? "is-active" : "")
            }
            onClick={() => {
              setActiveCategory(category);
              if (!expanded) setExpanded(true);
            }}
          >
            {category === "common" && <Star size={13} />}
            {(isEn ? categoryLabelsEn : categoryLabels)[category]}
          </button>
        ))}
        <button
          type="button"
          className="toolbar-collapse-button"
          onClick={() => setExpanded((value) => !value)}
          aria-expanded={expanded}
          title={
            expanded
              ? isEn
                ? "Collapse formula toolbar"
                : "收起公式选择栏"
              : isEn
                ? "Expand formula toolbar"
                : "展开公式选择栏"
          }
        >
          {expanded ? (
            <>
              {isEn ? "Collapse" : "收起"} <ChevronUp size={15} />
            </>
          ) : (
            <>
              {isEn ? "Expand" : "展开"} <ChevronDown size={15} />
            </>
          )}
        </button>
      </div>

      {expanded && (
        <div className="template-strip">
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
      )}
    </section>
  );
}
