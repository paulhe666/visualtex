import { useMemo, useState } from "react";
import { PanelLeftClose, Star } from "lucide-react";
import type { LatexCommand } from "../types/command";
import {
  categoryLabels,
  categoryLabelsEn,
  calculusCommandIds,
  commandRegistry,
  commonCommandIds,
} from "../autocomplete/commandRegistry";
import { MathPreview } from "../components/MathPreview";
import { useEditorStore } from "../stores/editorStore";

interface Props {
  onInsert: (command: LatexCommand) => void;
  onClose: () => void;
}

type MatrixDelimiter = "bmatrix" | "pmatrix" | "vmatrix";

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

const matrixSizes = Array.from({ length: 10 }, (_, index) => index + 1);
const matrixDelimiterOptions: Array<{
  id: MatrixDelimiter;
  preview: string;
  labelZh: string;
  labelEn: string;
}> = [
  {
    id: "vmatrix",
    preview: "\\begin{vmatrix}a&b\\\\c&d\\end{vmatrix}",
    labelZh: "竖线",
    labelEn: "Bars",
  },
  {
    id: "bmatrix",
    preview: "\\begin{bmatrix}a&b\\\\c&d\\end{bmatrix}",
    labelZh: "方括号",
    labelEn: "Brackets",
  },
  {
    id: "pmatrix",
    preview: "\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}",
    labelZh: "圆括号",
    labelEn: "Parentheses",
  },
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

function createMatrixCommand(
  rows: number,
  columns: number,
  delimiter: MatrixDelimiter,
): LatexCommand {
  const matrixBody = Array.from({ length: rows }, () =>
    Array.from({ length: columns }, () => "\\placeholder{}").join(" & "),
  ).join(" \\\\ ");
  const delimiterCopy = matrixDelimiterOptions.find(
    (option) => option.id === delimiter,
  ) ?? matrixDelimiterOptions[1];

  return {
    id: `custom-${delimiter}-${rows}x${columns}`,
    command: `\\begin{${delimiter}}`,
    insertTemplate: `\\begin{${delimiter}}${matrixBody}\\end{${delimiter}}`,
    previewLatex: delimiterCopy.preview,
    labelZh: `${rows}×${columns} ${delimiterCopy.labelZh}矩阵`,
    labelEn: `${rows}×${columns} ${delimiterCopy.labelEn.toLowerCase()} matrix`,
    aliases: ["matrix", delimiter],
    keywords: ["矩阵", "自定义矩阵", `${rows}x${columns}`],
    category: "matrix",
    defaultPriority: 120,
    supportedInMathMode: true,
  };
}

export function FormulaToolbar({ onInsert, onClose }: Props) {
  const [activeCategory, setActiveCategory] = useState("common");
  const [matrixRows, setMatrixRows] = useState(2);
  const [matrixColumns, setMatrixColumns] = useState(2);
  const [matrixDelimiter, setMatrixDelimiter] =
    useState<MatrixDelimiter>("bmatrix");
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";

  const visibleCommands = useMemo(() => {
    const preferredIds = activeCategory === "common"
      ? commonCommandIds
      : activeCategory === "calculus"
        ? calculusCommandIds
        : null;
    if (preferredIds) {
      return preferredIds
        .map((id) => commandRegistry.find((command) => command.id === id))
        .filter((command): command is LatexCommand => Boolean(command));
    }
    return commandRegistry.filter(
      (command) => command.category === activeCategory,
    );
  }, [activeCategory]);

  const insertCustomMatrix = () => {
    onInsert(createMatrixCommand(matrixRows, matrixColumns, matrixDelimiter));
  };

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
            data-category={category}
            aria-pressed={activeCategory === category}
            onClick={() => setActiveCategory(category)}
          >
            {category === "common" && <Star size={13} />}
            {(isEn ? categoryLabelsEn : categoryLabels)[category]}
          </button>
        ))}
      </nav>

      <div className="template-strip" aria-label={isEn ? "Formula templates" : "公式模板"}>
        {activeCategory === "matrix" && (
          <section className="matrix-builder" aria-label={isEn ? "Custom matrix" : "自定义矩阵"}>
            <div className="matrix-builder-heading">
              <div>
                <strong>{isEn ? "Custom matrix" : "自定义矩阵"}</strong>
                <span>{isEn ? "Up to 10 × 10" : "最大 10 × 10"}</span>
              </div>
              <span className="matrix-size-badge">
                {matrixRows} × {matrixColumns}
              </span>
            </div>

            <div className="matrix-delimiter-options" role="group" aria-label={isEn ? "Matrix delimiter" : "矩阵边界"}>
              {matrixDelimiterOptions.map((option) => (
                <button
                  key={option.id}
                  type="button"
                  className={matrixDelimiter === option.id ? "is-active" : ""}
                  aria-pressed={matrixDelimiter === option.id}
                  onClick={() => setMatrixDelimiter(option.id)}
                  title={isEn ? option.labelEn : option.labelZh}
                >
                  <MathPreview latex={option.preview} />
                  <span>{isEn ? option.labelEn : option.labelZh}</span>
                </button>
              ))}
            </div>

            <div className="matrix-dimension-row">
              <label>
                <span>{isEn ? "Rows" : "行数"}</span>
                <select
                  value={matrixRows}
                  onChange={(event) => setMatrixRows(Number(event.target.value))}
                >
                  {matrixSizes.map((size) => (
                    <option key={size} value={size}>{size}</option>
                  ))}
                </select>
              </label>
              <span aria-hidden="true">×</span>
              <label>
                <span>{isEn ? "Columns" : "列数"}</span>
                <select
                  value={matrixColumns}
                  onChange={(event) => setMatrixColumns(Number(event.target.value))}
                >
                  {matrixSizes.map((size) => (
                    <option key={size} value={size}>{size}</option>
                  ))}
                </select>
              </label>
            </div>

            <button
              type="button"
              className="matrix-insert-button"
              data-command-id="custom-matrix"
              onClick={insertCustomMatrix}
            >
              {isEn
                ? `Insert ${matrixRows} × ${matrixColumns} matrix`
                : `插入 ${matrixRows} × ${matrixColumns} 矩阵`}
            </button>
          </section>
        )}

        {visibleCommands.map((command) => (
          <button
            type="button"
            className={"template-button" + previewSizeClass(command)}
            data-command-id={command.id}
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
