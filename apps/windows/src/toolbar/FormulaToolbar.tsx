import { useMemo, useState } from "react";
import { Brackets, PanelLeftClose, Star } from "lucide-react";
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

const matrixGridCells = Array.from({ length: 100 }, (_, index) => ({
  row: Math.floor(index / 10) + 1,
  column: (index % 10) + 1,
}));
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
  const [matrixHover, setMatrixHover] = useState<{
    rows: number;
    columns: number;
  } | null>(null);
  const [matrixDelimiter, setMatrixDelimiter] =
    useState<MatrixDelimiter>("bmatrix");
  const language = useEditorStore((state) => state.language);
  const autoPairDelimiters = useEditorStore(
    (state) => state.autoPairDelimiters,
  );
  const setAutoPairDelimiters = useEditorStore(
    (state) => state.setAutoPairDelimiters,
  );
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
  const previewRows = matrixHover?.rows ?? matrixRows;
  const previewColumns = matrixHover?.columns ?? matrixColumns;

  return (
    <aside
      className="formula-toolbar"
      aria-label={isEn ? "Formula toolbar" : "公式工具栏"}
    >
      <header className="formula-toolbar-header">
        <strong>{isEn ? "Formula tools" : "公式工具"}</strong>
        <div className="formula-toolbar-actions">
          <button
            type="button"
            className={
              "icon-button compact" +
              (autoPairDelimiters ? " is-active" : "")
            }
            aria-pressed={autoPairDelimiters}
            onClick={() => setAutoPairDelimiters(!autoPairDelimiters)}
            aria-label={isEn ? "Auto-pair delimiters" : "自动补全成对符号"}
            title={
              isEn
                ? "Auto-pair brackets, braces and vertical bars"
                : "自动补全括号、花括号和竖线"
            }
          >
            <Brackets size={16} />
          </button>
          <button
            type="button"
            className="icon-button compact"
            onClick={onClose}
            aria-label={isEn ? "Hide formula tools" : "隐藏公式工具"}
            title={isEn ? "Hide formula tools" : "隐藏公式工具"}
          >
            <PanelLeftClose size={16} />
          </button>
        </div>
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
              <span className="matrix-size-badge" aria-live="polite">
                {previewRows} × {previewColumns}
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

            <div className="matrix-size-picker">
              <span className="matrix-size-picker-label">
                {isEn
                  ? "Move to preview · Click to select rows and columns"
                  : "移动预览 · 单击选择行数和列数"}
              </span>
              <div
                className="matrix-size-grid"
                role="grid"
                aria-label={
                  isEn
                    ? "Select matrix rows and columns"
                    : "选择矩阵行数和列数"
                }
                aria-rowcount={10}
                aria-colcount={10}
                onPointerLeave={() => setMatrixHover(null)}
                onBlur={(event) => {
                  if (
                    !event.currentTarget.contains(
                      event.relatedTarget as Node | null,
                    )
                  ) {
                    setMatrixHover(null);
                  }
                }}
              >
                {matrixGridCells.map(({ row, column }) => {
                  const previewed =
                    row <= previewRows && column <= previewColumns;
                  const selectedCorner =
                    row === matrixRows && column === matrixColumns;
                  return (
                    <button
                      key={`${row}-${column}`}
                      type="button"
                      role="gridcell"
                      className={
                        "matrix-size-cell" +
                        (previewed ? " is-previewed" : "") +
                        (selectedCorner ? " is-selected-corner" : "")
                      }
                      aria-label={
                        isEn
                          ? `${row} rows by ${column} columns`
                          : `${row} 行 ${column} 列`
                      }
                      aria-selected={selectedCorner}
                      data-matrix-rows={row}
                      data-matrix-columns={column}
                      onPointerEnter={() =>
                        setMatrixHover({ rows: row, columns: column })
                      }
                      onFocus={() =>
                        setMatrixHover({ rows: row, columns: column })
                      }
                      onClick={() => {
                        setMatrixRows(row);
                        setMatrixColumns(column);
                        setMatrixHover(null);
                      }}
                    />
                  );
                })}
              </div>
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
