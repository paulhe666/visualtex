import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
} from "react";
import { createPortal } from "react-dom";
import { MathfieldElement, convertLatexToMarkup } from "mathlive";
import {
  categories,
  categoryLabels,
  commandsForCategory,
  createMatrixCommand,
  searchCommands,
  templateForSelection,
  type CommandCategory,
  type LatexCommand,
} from "./commands";
import "mathlive/fonts.css";
import "mathlive/static.css";
import "./styles.css";

export interface MathNodeEditorProps {
  value: string;
  disabled?: boolean;
  autoFocus?: boolean;
  showToolbar?: boolean;
  showCandidates?: boolean;
  onChange: (latex: string) => void;
  onCommit?: (latex: string) => void;
}

function MathPreview({ latex }: { latex: string }) {
  const markup = useMemo(
    () => convertLatexToMarkup(latex, { defaultMode: "math" }),
    [latex],
  );
  return (
    <span
      className="vt-formula-preview"
      aria-hidden="true"
      dangerouslySetInnerHTML={{ __html: markup }}
    />
  );
}

function findTrailingCommandRange(
  field: MathfieldElement,
  query: string,
): [number, number] | null {
  const expected = query.trim();
  if (!expected) return null;
  const ends = Array.from(new Set([field.position, field.lastOffset].filter((offset) => offset >= 0)));
  for (const end of ends) {
    for (let start = end; start >= 0; start -= 1) {
      if (field.getValue(start, end, "latex").trim() === expected) return [start, end];
    }
  }
  return null;
}

function trailingCommand(field: MathfieldElement): string | null {
  const prefix = field.getValue(0, field.position, "latex");
  return prefix.match(/\\[\p{L}]*$/u)?.[0] ?? null;
}

function FormulaToolbar({
  onInsert,
}: {
  onInsert: (command: LatexCommand) => void;
}) {
  const [activeCategory, setActiveCategory] = useState<CommandCategory>("common");
  const [rows, setRows] = useState(2);
  const [columns, setColumns] = useState(2);
  const [delimiter, setDelimiter] = useState<"bmatrix" | "pmatrix" | "vmatrix">("bmatrix");
  const commands = useMemo(() => commandsForCategory(activeCategory), [activeCategory]);

  return (
    <section className="vt-formula-toolbar" aria-label="公式工具">
      <nav className="vt-formula-toolbar-tabs" aria-label="公式分类">
        {categories.map((category) => (
          <button
            type="button"
            key={category}
            className={activeCategory === category ? "active" : ""}
            aria-pressed={activeCategory === category}
            onClick={() => setActiveCategory(category)}
          >
            {categoryLabels[category]}
          </button>
        ))}
      </nav>
      {activeCategory === "matrix" && (
        <div className="vt-matrix-builder">
          <strong>自定义矩阵</strong>
          <label>
            <span>行</span>
            <select value={rows} onChange={(event) => setRows(Number(event.target.value))}>
              {Array.from({ length: 10 }, (_, index) => index + 1).map((value) => (
                <option value={value} key={value}>{value}</option>
              ))}
            </select>
          </label>
          <span>×</span>
          <label>
            <span>列</span>
            <select value={columns} onChange={(event) => setColumns(Number(event.target.value))}>
              {Array.from({ length: 10 }, (_, index) => index + 1).map((value) => (
                <option value={value} key={value}>{value}</option>
              ))}
            </select>
          </label>
          <select value={delimiter} onChange={(event) => setDelimiter(event.target.value as typeof delimiter)}>
            <option value="bmatrix">方括号</option>
            <option value="pmatrix">圆括号</option>
            <option value="vmatrix">行列式</option>
          </select>
          <button type="button" onClick={() => onInsert(createMatrixCommand(rows, columns, delimiter))}>
            插入 {rows}×{columns}
          </button>
        </div>
      )}
      <div className="vt-formula-command-grid">
        {commands.map((candidate) => (
          <button
            type="button"
            key={candidate.id}
            title={`${candidate.labelZh} · ${candidate.command}`}
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => onInsert(candidate)}
          >
            <MathPreview latex={candidate.previewLatex} />
            <span>{candidate.labelZh}</span>
          </button>
        ))}
      </div>
    </section>
  );
}

export function MathNodeEditor({
  value,
  disabled = false,
  autoFocus = false,
  showToolbar = false,
  showCandidates = true,
  onChange,
  onCommit,
}: MathNodeEditorProps) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const fieldRef = useRef<MathfieldElement | null>(null);
  const suppressRef = useRef(false);
  const callbacksRef = useRef({ onChange, onCommit });
  callbacksRef.current = { onChange, onCommit };
  const [query, setQuery] = useState<string | null>(null);
  const [selectedSuggestion, setSelectedSuggestion] = useState(0);
  const [candidatePosition, setCandidatePosition] = useState({
    left: 12,
    top: 12,
    width: 420,
    ready: false,
  });
  const suggestions = useMemo(
    () => (showCandidates && query !== null ? searchCommands(query, 8) : []),
    [query, showCandidates],
  );
  const suggestionsRef = useRef(suggestions);
  const selectedSuggestionRef = useRef(selectedSuggestion);
  const queryRef = useRef(query);
  suggestionsRef.current = suggestions;
  selectedSuggestionRef.current = selectedSuggestion;
  queryRef.current = query;

  const publishValue = useCallback((field: MathfieldElement) => {
    if (!suppressRef.current) callbacksRef.current.onChange(field.value);
    const nextQuery = trailingCommand(field);
    setQuery(nextQuery);
    setSelectedSuggestion(0);
  }, []);

  const insertCommand = useCallback((candidate: LatexCommand, replaceTypedQuery = false) => {
    const field = fieldRef.current;
    if (!field || field.readOnly) return;
    let selectedLatex = field.selectionIsCollapsed ? "" : field.getValue(field.selection);
    if (replaceTypedQuery && queryRef.current) {
      const range = findTrailingCommandRange(field, queryRef.current);
      if (range) {
        field.selection = { ranges: [range], direction: "none" };
        selectedLatex = "";
      }
    }
    field.insert(templateForSelection(candidate, selectedLatex), {
      insertionMode: "replaceSelection",
      selectionMode: "placeholder",
      format: "latex",
    });
    setQuery(null);
    setSelectedSuggestion(0);
    callbacksRef.current.onChange(field.value);
    queueMicrotask(() => field.focus());
  }, []);

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const field = new MathfieldElement();
    field.value = value;
    field.smartFence = true;
    field.smartMode = false;
    field.popoverPolicy = "off";
    field.maxMatrixCols = 10;
    field.mathVirtualKeyboardPolicy = "manual";
    field.className = "vt-math-field vt-visualtex-math-field";
    field.readOnly = disabled;
    field.setAttribute("aria-label", "可视化 LaTeX 公式编辑器");

    const handleInput = () => publishValue(field);
    const handleChange = () => callbacksRef.current.onCommit?.(field.value);
    const handleKeyDown = (event: KeyboardEvent) => {
      const currentSuggestions = suggestionsRef.current;
      if (currentSuggestions.length > 0) {
        if (event.key === "ArrowDown" || event.key === "ArrowUp") {
          event.preventDefault();
          event.stopPropagation();
          const direction = event.key === "ArrowDown" ? 1 : -1;
          setSelectedSuggestion((current) =>
            (current + direction + currentSuggestions.length) % currentSuggestions.length,
          );
          return;
        }
        if (event.key === "Enter" || event.key === "Tab") {
          event.preventDefault();
          event.stopPropagation();
          const candidate = currentSuggestions[selectedSuggestionRef.current] ?? currentSuggestions[0];
          if (candidate) insertCommand(candidate, true);
          return;
        }
        if (event.key === "Escape") {
          event.preventDefault();
          event.stopPropagation();
          setQuery(null);
          return;
        }
      }
      if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
        event.preventDefault();
        callbacksRef.current.onCommit?.(field.value);
      }
    };

    field.addEventListener("input", handleInput);
    field.addEventListener("change", handleChange);
    field.addEventListener("keydown", handleKeyDown);
    host.appendChild(field);
    fieldRef.current = field;
    if (autoFocus && !disabled) queueMicrotask(() => field.focus());

    return () => {
      field.removeEventListener("input", handleInput);
      field.removeEventListener("change", handleChange);
      field.removeEventListener("keydown", handleKeyDown);
      field.remove();
      fieldRef.current = null;
    };
  }, []);

  useEffect(() => {
    const field = fieldRef.current;
    if (!field || field.value === value) return;
    suppressRef.current = true;
    field.value = value;
    suppressRef.current = false;
  }, [value]);

  useEffect(() => {
    if (fieldRef.current) fieldRef.current.readOnly = disabled;
  }, [disabled]);

  useLayoutEffect(() => {
    if (suggestions.length === 0) return;
    const update = () => {
      const host = hostRef.current;
      if (!host) return;
      const bounds = host.getBoundingClientRect();
      const width = Math.max(300, Math.min(460, window.innerWidth - 24));
      const estimatedHeight = Math.min(390, 38 + suggestions.length * 55 + 14);
      const left = Math.max(12, Math.min(bounds.left, window.innerWidth - width - 12));
      const below = bounds.bottom + 7;
      const above = bounds.top - estimatedHeight - 7;
      const top = below + estimatedHeight <= window.innerHeight - 12
        ? below
        : above >= 12
          ? above
          : Math.max(12, window.innerHeight - estimatedHeight - 12);
      setCandidatePosition({ left, top, width, ready: true });
    };
    update();
    window.addEventListener("resize", update);
    window.addEventListener("scroll", update, true);
    return () => {
      window.removeEventListener("resize", update);
      window.removeEventListener("scroll", update, true);
    };
  }, [suggestions.length]);

  const handleEditorKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.key === "Escape" && query !== null) {
      event.stopPropagation();
      setQuery(null);
    }
  };

  return (
    <div className={`vt-math-editor${showToolbar ? " with-toolbar" : ""}`} onKeyDown={handleEditorKeyDown}>
      <div className="vt-math-field-shell">
        <div ref={hostRef} className="vt-math-field-host" />
      </div>
      {suggestions.length > 0 && createPortal(
        <div
          className="vt-command-suggestions"
          role="listbox"
          aria-label="LaTeX 命令候选"
          style={{
            position: "fixed",
            left: candidatePosition.left,
            top: candidatePosition.top,
            width: candidatePosition.width,
            visibility: candidatePosition.ready ? "visible" : "hidden",
          }}
        >
          <header><span>命令候选</span><small>↑↓ 选择 · Enter 插入</small></header>
          {suggestions.map((candidate, index) => (
            <button
              type="button"
              role="option"
              aria-selected={index === selectedSuggestion}
              className={index === selectedSuggestion ? "active" : ""}
              key={candidate.id}
              onMouseEnter={() => setSelectedSuggestion(index)}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => insertCommand(candidate, true)}
            >
              <MathPreview latex={candidate.previewLatex} />
              <span><strong>{candidate.command}</strong><small>{candidate.labelZh}</small></span>
            </button>
          ))}
        </div>,
        document.body,
      )}
      {showToolbar && <FormulaToolbar onInsert={(candidate) => insertCommand(candidate, false)} />}
    </div>
  );
}

export { commandRegistry, searchCommands, templateForSelection } from "./commands";
