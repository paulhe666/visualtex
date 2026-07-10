import {
  useEffect,
  useImperativeHandle,
  useMemo,
  useRef,
  useState,
  forwardRef,
} from "react";
import { MathfieldElement } from "mathlive";
import { Plus } from "lucide-react";
import type { CommandSource, LatexCommand } from "../types/command";
import { searchCommands } from "../autocomplete/CommandSearchEngine";
import { CommandSuggestionPopup } from "../autocomplete/CommandSuggestionPopup";
import { useEditorStore } from "../stores/editorStore";
import { normalizeChineseLatex } from "./normalizeChineseLatex";

export interface MathEditorHandle {
  insertCommand: (command: LatexCommand, source?: "toolbar" | "history" | "shortcut") => void;
  undo: () => void;
  redo: () => void;
  focus: () => void;
  addLine: () => void;
}

interface Props {
  lines: string[];
  onChange: (lines: string[]) => void;
  zoom: number;
}

interface FormulaFieldProps {
  lineId: string;
  index: number;
  latex: string;
  zoom: number;
  language: "cn" | "en";
  register: (lineId: string, field: MathfieldElement | null) => void;
  onInput: (index: number, field: MathfieldElement) => void;
  onFocus: (index: number, field: MathfieldElement) => void;
  onKeyDown: (index: number, event: KeyboardEvent, field: MathfieldElement) => void;
}

const trailingCommand = /\\([\p{L}]*)$/u;

function templateForSelection(
  command: LatexCommand,
  selectedLatex: string,
): string {
  if (!selectedLatex) return command.insertTemplate;

  switch (command.id) {
    case "scripts":
      return selectedLatex + "_{\\placeholder{}}^{\\placeholder{}}";
    case "lower-script":
      return selectedLatex + "_{\\placeholder{}}";
    case "upper-script":
      return selectedLatex + "^{\\placeholder{}}";
    case "sum":
    case "series":
      return "\\sum_{\\placeholder{}}^{\\placeholder{}} " + selectedLatex;
    case "prod":
    case "productseries":
      return "\\prod_{\\placeholder{}}^{\\placeholder{}} " + selectedLatex;
    case "int":
      return "\\int_{\\placeholder{}}^{\\placeholder{}} " + selectedLatex + "\\,\\mathrm{d}\\placeholder{}";
    case "intplain":
      return "\\int " + selectedLatex + "\\,\\mathrm{d}\\placeholder{}";
    case "frac":
    case "smallfrac":
    case "displayfrac":
      return command.command + "{" + selectedLatex + "}{\\placeholder{}}";
    case "sqrt":
      return "\\sqrt{" + selectedLatex + "}";
    case "parentheses":
      return "\\left(" + selectedLatex + "\\right)";
    case "brackets":
      return "\\left[" + selectedLatex + "\\right]";
    case "braces":
      return "\\left\\{" + selectedLatex + "\\right\\}";
    case "absolute":
      return "\\left|" + selectedLatex + "\\right|";
    default:
      return command.insertTemplate.replace("\\placeholder{}", selectedLatex);
  }
}

function FormulaField(props: FormulaFieldProps) {
  const hostRef = useRef<HTMLDivElement>(null);
  const fieldRef = useRef<MathfieldElement | null>(null);
  const propsRef = useRef(props);
  propsRef.current = props;

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;

    const lineId = propsRef.current.lineId;
    const field = new MathfieldElement();
    field.value = propsRef.current.latex;
    field.className = "visual-mathfield";
    field.smartMode = false;
    field.setAttribute("math-virtual-keyboard-policy", "manual");
    const isEn = propsRef.current.language === "en";
    field.setAttribute(
      "aria-label",
      isEn
        ? "Formula line " + (propsRef.current.index + 1)
        : "第 " + (propsRef.current.index + 1) + " 行公式",
    );
    field.setAttribute(
      "placeholder",
      propsRef.current.index === 0
        ? isEn
          ? "\\text{Enter a formula, press Enter for a new line}"
          : "\\text{输入公式，按 Enter 新建下一行}"
        : isEn
          ? "\\text{Continue entering a formula}"
          : "\\text{继续输入公式}",
    );
    field.style.fontSize = 34 * propsRef.current.zoom + "px";

    let composing = false;
    const handleCompositionStart = () => {
      composing = true;
    };
    const handleCompositionEnd = () => {
      composing = false;
      propsRef.current.onInput(propsRef.current.index, field);
    };
    const handleInput = () => {
      if (!composing) propsRef.current.onInput(propsRef.current.index, field);
    };
    const handleFocus = () => {
      propsRef.current.onFocus(propsRef.current.index, field);
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.isComposing || composing) return;
      propsRef.current.onKeyDown(propsRef.current.index, event, field);
    };
    const handlePointerDown = (event: PointerEvent) => {
      const offset = field.getOffsetFromPoint(event.clientX, event.clientY, {
        bias: 1,
      });
      if (offset !== field.lastOffset) return;

      event.preventDefault();
      field.focus();
      field.position = field.lastOffset;
      propsRef.current.onFocus(propsRef.current.index, field);
    };

    host.replaceChildren(field);
    fieldRef.current = field;
    propsRef.current.register(lineId, field);
    field.addEventListener("compositionstart", handleCompositionStart);
    field.addEventListener("compositionend", handleCompositionEnd);
    field.addEventListener("input", handleInput);
    field.addEventListener("focus", handleFocus);
    field.addEventListener("keydown", handleKeyDown, true);
    field.addEventListener("pointerdown", handlePointerDown, true);

    return () => {
      field.removeEventListener("compositionstart", handleCompositionStart);
      field.removeEventListener("compositionend", handleCompositionEnd);
      field.removeEventListener("input", handleInput);
      field.removeEventListener("focus", handleFocus);
      field.removeEventListener("keydown", handleKeyDown, true);
      field.removeEventListener("pointerdown", handlePointerDown, true);
      propsRef.current.register(lineId, null);
      fieldRef.current = null;
      host.replaceChildren();
    };
  }, []);

  useEffect(() => {
    const field = fieldRef.current;
    if (!field || field.value === props.latex) return;

    // 本地输入仅因中文规范化而与 store 不同时，不重建 MathLive 模型，
    // 否则会丢失当前光标、选区以及删除键的内部状态。
    if (normalizeChineseLatex(field.value) === props.latex) return;
    field.value = props.latex;
  }, [props.latex]);

  useEffect(() => {
    if (fieldRef.current) {
      fieldRef.current.style.fontSize = 34 * props.zoom + "px";
    }
  }, [props.zoom]);

  useEffect(() => {
    const field = fieldRef.current;
    if (!field) return;
    const isEn = props.language === "en";
    field.setAttribute(
      "aria-label",
      isEn
        ? "Formula line " + (props.index + 1)
        : "第 " + (props.index + 1) + " 行公式",
    );
    field.setAttribute(
      "placeholder",
      props.index === 0
        ? isEn
          ? "\\text{Enter a formula, press Enter for a new line}"
          : "\\text{输入公式，按 Enter 新建下一行}"
        : isEn
          ? "\\text{Continue entering a formula}"
          : "\\text{继续输入公式}",
    );
  }, [props.index, props.language]);

  return <div ref={hostRef} className="mathfield-host" />;
}

export const MathEditor = forwardRef<MathEditorHandle, Props>(
  function MathEditor({ lines, onChange, zoom }, ref) {
    const surfaceRef = useRef<HTMLDivElement>(null);
    const fieldRefs = useRef(new Map<string, MathfieldElement>());
    const lineIdsRef = useRef<string[]>([]);
    const linesRef = useRef(lines);
    const activeIndexRef = useRef(0);
    const activeLineIdRef = useRef<string | null>(null);
    const pendingFocusRef = useRef<{
      lineId: string;
      index: number;
      moveToEnd: boolean;
    } | null>(null);
    const [activeIndex, setActiveIndex] = useState(0);
    const [query, setQuery] = useState("");
    const [selectedIndex, setSelectedIndex] = useState(0);
    const [popupPosition, setPopupPosition] = useState({ left: 72, top: 132 });
    const selectedIndexRef = useRef(0);
    const queryRef = useRef("");
    const usage = useEditorStore((state) => state.usage);
    const personalize = useEditorStore((state) => state.personalize);
    const suggestionCount = useEditorStore((state) => state.suggestionCount);
    const recordCommand = useEditorStore((state) => state.recordCommand);
    const language = useEditorStore((state) => state.language);
    const isEn = language === "en";

    linesRef.current = lines;
    while (lineIdsRef.current.length < lines.length) {
      lineIdsRef.current.push(crypto.randomUUID());
    }
    if (lineIdsRef.current.length > lines.length) {
      lineIdsRef.current.length = lines.length;
    }
    if (!activeLineIdRef.current && lineIdsRef.current.length) {
      activeLineIdRef.current = lineIdsRef.current[0];
    }

    const suggestions = useMemo(
      () =>
        query
          ? searchCommands(query, usage, personalize, suggestionCount)
          : [],
      [query, usage, personalize, suggestionCount],
    );

    const focusLine = (
      index: number,
      moveToEnd = false,
      remainingAttempts = 8,
    ) => {
      const lineId = lineIdsRef.current[index];
      if (!lineId) return;

      window.requestAnimationFrame(() => {
        const field = fieldRefs.current.get(lineId);
        if (!field?.isConnected) {
          if (remainingAttempts > 0) {
            window.setTimeout(
              () => focusLine(index, moveToEnd, remainingAttempts - 1),
              0,
            );
          }
          return;
        }

        activeIndexRef.current = index;
        activeLineIdRef.current = lineId;
        setActiveIndex(index);
        field.focus();
        if (moveToEnd) {
          field.executeCommand("moveToMathfieldEnd");
          field.position = field.lastOffset;
        }
        pendingFocusRef.current = null;

        if (moveToEnd) {
          window.requestAnimationFrame(() => {
            if (!field.isConnected) return;
            field.focus();
            field.executeCommand("moveToMathfieldEnd");
            field.position = field.lastOffset;
          });
        }
      });
    };

    const registerField = (
      lineId: string,
      field: MathfieldElement | null,
    ) => {
      if (field) fieldRefs.current.set(lineId, field);
      else fieldRefs.current.delete(lineId);

      const pending = pendingFocusRef.current;
      if (field && pending?.lineId === lineId) {
        focusLine(pending.index, pending.moveToEnd);
      }
    };

    const updatePopupPosition = (field: MathfieldElement) => {
      const fieldWithCaret = field as MathfieldElement & {
        getCaretPoint?: () => { x: number; y: number; height?: number };
      };
      const surface = surfaceRef.current;
      if (!surface) return;
      const surfaceRect = surface.getBoundingClientRect();
      const caret = fieldWithCaret.getCaretPoint?.();

      if (caret) {
        setPopupPosition({
          left: Math.max(20, Math.min(surfaceRect.width - 370, caret.x - surfaceRect.left)),
          top: Math.max(
            64,
            caret.y - surfaceRect.top + (caret.height ?? 28) + 8,
          ),
        });
      } else {
        const fieldRect = field.getBoundingClientRect();
        setPopupPosition({
          left: Math.max(20, fieldRect.left - surfaceRect.left + 48),
          top: fieldRect.bottom - surfaceRect.top + 8,
        });
      }
    };

    const syncField = (index: number, field: MathfieldElement) => {
      const normalized = normalizeChineseLatex(field.value);
      const nextLines = [...linesRef.current];
      nextLines[index] = normalized;
      linesRef.current = nextLines;
      onChange(nextLines);

      const match = normalized.match(trailingCommand);
      if (match) {
        setQuery("\\" + match[1]);
        setSelectedIndex(0);
        requestAnimationFrame(() => updatePopupPosition(field));
      } else {
        setQuery("");
      }
    };

    const insertCommand = (
      command: LatexCommand,
      source: CommandSource = "toolbar",
      activeQuery = "",
    ) => {
      let targetIndex = lineIdsRef.current.indexOf(
        activeLineIdRef.current ?? "",
      );
      if (targetIndex < 0) {
        targetIndex = Math.min(
          activeIndexRef.current,
          Math.max(0, linesRef.current.length - 1),
        );
      }

      let targetLineId = lineIdsRef.current[targetIndex];
      let field = targetLineId
        ? fieldRefs.current.get(targetLineId)
        : undefined;

      if (!field?.isConnected) {
        targetIndex = lineIdsRef.current.findIndex((lineId) =>
          Boolean(fieldRefs.current.get(lineId)?.isConnected),
        );
        targetLineId = lineIdsRef.current[targetIndex];
        field = targetLineId
          ? fieldRefs.current.get(targetLineId)
          : undefined;
      }
      if (!field?.isConnected || !targetLineId) return;

      activeIndexRef.current = targetIndex;
      activeLineIdRef.current = targetLineId;
      setActiveIndex(targetIndex);
      field.focus();

      if (activeQuery) {
        for (let index = 0; index < activeQuery.length; index += 1) {
          field.executeCommand("deleteBackward");
        }
      }

      const selectedLatex = field.selectionIsCollapsed
        ? ""
        : field.getValue(field.selection);
      const insertionTemplate = templateForSelection(command, selectedLatex);

      const finishInsertion = () => {
        recordCommand(command.id, activeQuery, source);
        setQuery("");
        setSelectedIndex(0);
        syncField(targetIndex, field);
        field.focus();
      };
      const tryInsert = () =>
        field.insert(insertionTemplate, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceSelection",
          selectionMode: "placeholder",
          focus: true,
          scrollIntoView: false,
        });

      if (tryInsert()) {
        finishInsertion();
        return;
      }

      // WebKit 偶尔会在焦点刚恢复时拒绝首次插入，下一帧重试一次。
      window.requestAnimationFrame(() => {
        if (!field.isConnected) return;
        field.focus();
        if (tryInsert()) finishInsertion();
      });
    };

    const addLineAfter = (index: number) => {
      const nextLines = [...linesRef.current];
      const nextIndex = index + 1;
      const nextLineId = crypto.randomUUID();
      nextLines.splice(nextIndex, 0, "");
      lineIdsRef.current.splice(nextIndex, 0, nextLineId);
      linesRef.current = nextLines;
      pendingFocusRef.current = {
        lineId: nextLineId,
        index: nextIndex,
        moveToEnd: false,
      };
      onChange(nextLines);
      setQuery("");
      window.setTimeout(() => focusLine(nextIndex), 0);
    };

    const removeEmptyLine = (index: number) => {
      if (linesRef.current.length <= 1) return;

      const removedLineId = lineIdsRef.current[index];
      const nextLines = linesRef.current.filter((_, lineIndex) => lineIndex !== index);
      lineIdsRef.current.splice(index, 1);
      if (removedLineId) fieldRefs.current.delete(removedLineId);
      linesRef.current = nextLines;

      const previousIndex = Math.max(0, index - 1);
      const previousLineId = lineIdsRef.current[previousIndex];
      activeIndexRef.current = previousIndex;
      activeLineIdRef.current = previousLineId;
      setActiveIndex(previousIndex);
      pendingFocusRef.current = {
        lineId: previousLineId,
        index: previousIndex,
        moveToEnd: true,
      };
      onChange(nextLines);
      setQuery("");
      window.setTimeout(() => focusLine(previousIndex, true), 0);
    };

    const handleKeyDown = (
      index: number,
      lineId: string,
      event: KeyboardEvent,
      field: MathfieldElement,
    ) => {
      activeIndexRef.current = index;
      activeLineIdRef.current = lineId;
      setActiveIndex(index);

      const liveMatch = field.value.match(trailingCommand);
      const liveQuery = liveMatch ? "\\" + liveMatch[1] : "";
      const liveSuggestions = liveQuery
        ? searchCommands(
            liveQuery,
            useEditorStore.getState().usage,
            useEditorStore.getState().personalize,
            useEditorStore.getState().suggestionCount,
          )
        : [];

      if (liveSuggestions.length) {
        if (event.key === "ArrowDown") {
          event.preventDefault();
          event.stopPropagation();
          setSelectedIndex((current) => (current + 1) % liveSuggestions.length);
          return;
        }
        if (event.key === "ArrowUp") {
          event.preventDefault();
          event.stopPropagation();
          setSelectedIndex(
            (current) => (current - 1 + liveSuggestions.length) % liveSuggestions.length,
          );
          return;
        }
        if (event.key === "Enter" || event.key === "Tab") {
          event.preventDefault();
          event.stopPropagation();
          const suggestionIndex = Math.min(
            selectedIndexRef.current,
            liveSuggestions.length - 1,
          );
          insertCommand(
            liveSuggestions[suggestionIndex],
            "candidate",
            liveQuery,
          );
          return;
        }
        if (event.key === "Escape") {
          event.preventDefault();
          setQuery("");
          return;
        }
      }

      if (event.key === "Enter") {
        event.preventDefault();
        event.stopPropagation();
        syncField(index, field);
        addLineAfter(index);
        return;
      }

      if (
        event.key === "Backspace" &&
        field.value.trim() === "" &&
        linesRef.current.length > 1
      ) {
        event.preventDefault();
        event.stopPropagation();
        removeEmptyLine(index);
      }
    };

    useImperativeHandle(ref, () => ({
      insertCommand,
      undo: () => {
        const index = activeIndexRef.current;
        const lineId = activeLineIdRef.current;
        const field = lineId ? fieldRefs.current.get(lineId) : undefined;
        field?.executeCommand("undo");
        if (field) requestAnimationFrame(() => syncField(index, field));
      },
      redo: () => {
        const index = activeIndexRef.current;
        const lineId = activeLineIdRef.current;
        const field = lineId ? fieldRefs.current.get(lineId) : undefined;
        field?.executeCommand("redo");
        if (field) requestAnimationFrame(() => syncField(index, field));
      },
      focus: () => focusLine(activeIndexRef.current),
      addLine: () => addLineAfter(linesRef.current.length - 1),
    }));

    useEffect(() => {
      selectedIndexRef.current = selectedIndex;
      if (selectedIndex >= suggestions.length) setSelectedIndex(0);
    }, [suggestions.length, selectedIndex]);

    useEffect(() => {
      queryRef.current = query;
    }, [query]);

    useEffect(() => {
      if (
        activeIndex >= lines.length ||
        !lineIdsRef.current.includes(activeLineIdRef.current ?? "")
      ) {
        const nextActive = Math.max(0, lines.length - 1);
        activeIndexRef.current = nextActive;
        activeLineIdRef.current = lineIdsRef.current[nextActive] ?? null;
        setActiveIndex(nextActive);
      }
    }, [lines.length, activeIndex]);

    return (
      <div ref={surfaceRef} className="editor-surface multi-line-editor">
        <div className="mathfield-stack">
          {lines.map((line, index) => {
            const lineId = lineIdsRef.current[index];
            return (
            <div
              className={"formula-line " + (index === activeIndex ? "is-active" : "")}
              key={lineId}
            >
              <span className="formula-line-number">
                {String(index + 1).padStart(2, "0")}
              </span>
              <FormulaField
                lineId={lineId}
                index={index}
                latex={line}
                zoom={zoom}
                language={language}
                register={registerField}
                onInput={syncField}
                onFocus={(lineIndex, field) => {
                  activeIndexRef.current = lineIndex;
                  activeLineIdRef.current = lineId;
                  setActiveIndex(lineIndex);
                  const match = field.value.match(trailingCommand);
                  if (match) {
                    setQuery("\\" + match[1]);
                    requestAnimationFrame(() => updatePopupPosition(field));
                  }
                }}
                onKeyDown={(lineIndex, event, field) =>
                  handleKeyDown(lineIndex, lineId, event, field)
                }
              />
            </div>
            );
          })}

          <button
            type="button"
            className="add-formula-line"
            onClick={() => addLineAfter(linesRef.current.length - 1)}
          >
            <Plus size={15} />
            {isEn ? "Add formula line" : "添加公式行"}
          </button>
        </div>

        <p className="editor-help">
          {isEn
            ? "Enter: new line · Backspace on empty line: delete · Tab: next placeholder · ↑↓: commands"
            : "Enter 新建下一行 · 空行按 Backspace 删除 · Tab 切换占位符 · ↑↓ 选择命令"}
        </p>
        <CommandSuggestionPopup
          suggestions={suggestions}
          selectedIndex={selectedIndex}
          position={popupPosition}
          usage={usage}
          onSelect={(command) =>
            insertCommand(command, "candidate", queryRef.current)
          }
        />
      </div>
    );
  },
);
