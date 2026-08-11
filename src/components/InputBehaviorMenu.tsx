import {
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
} from "react";
import { createPortal } from "react-dom";
import { convertVisualTexLatexToMarkup } from "../editor/mathLiveIntegralCompatibility";
import {
  ArrowLeft,
  ArrowRight,
  ChevronDown,
  MousePointerClick,
  Search,
} from "lucide-react";
import type { InputBehaviorSettingKey } from "../types/formula";
import { useEditorStore } from "../stores/editorStore";
import {
  visualTexAutoEscapeInlineShortcuts,
  visualTexAutoEscapeShortcutGroups,
  type VisualTexInlineShortcutDefinition,
  type VisualTexInlineShortcutDefinitions,
} from "../editor/normalizeChineseLatex";

interface InputBehaviorOption {
  key: InputBehaviorSettingKey;
  titleZh: string;
  titleEn: string;
  descriptionZh: string;
  descriptionEn: string;
}

const AUTO_ESCAPE_OPTIONS: InputBehaviorOption[] = [
  {
    key: "autoEscapeShortcuts",
    titleZh: "常用数学快捷转义",
    titleEn: "Common math shortcuts",
    descriptionZh: "控制 alpha、>=、hat 等快捷映射；微分元、函数名等正体自动检测独立运行，不受此开关影响",
    descriptionEn: "Controls shortcuts such as alpha, >= and hat; upright detection for differentials and function names remains independent",
  },
];

const CARET_BEHAVIOR_OPTIONS: InputBehaviorOption[] = [
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
  {
    key: "autoExitWrapperCommand",
    titleZh: "字体命令输入后跳出",
    titleEn: "Exit font command after input",
    descriptionZh: "开启时输入一个字符后自动结束；关闭后可连续输入，并按 Enter 确认 mathbb、mathbf、mathcal 等字体作用域",
    descriptionEn: "When enabled, exit after one character; when disabled, keep typing and press Enter to confirm mathbb, mathbf, mathcal and similar font scopes",
  },
];

const COMMAND_SUGGESTION_OPTIONS: InputBehaviorOption[] = [
  {
    key: "showStructuredCommandSuggestions",
    titleZh: "求和、积分等结构候选框",
    titleEn: "Structured command suggestions",
    descriptionZh: "控制 VisualTeX 的大型候选框，默认开启；不影响 MathLive 原生命令提示框",
    descriptionEn: "Controls the large VisualTeX panel for sums, integrals and similar structures; does not affect MathLive's native command panel",
  },
  {
    key: "showOtherCommandSuggestions",
    titleZh: "其他命令候选框",
    titleEn: "Other command suggestions",
    descriptionZh: "控制除求和、积分等结构外的 VisualTeX 大型候选框，默认关闭",
    descriptionEn: "Controls the large VisualTeX panel for commands other than sums, integrals and similar structures; off by default",
  },
];

function shortcutLatex(definition: VisualTexInlineShortcutDefinition) {
  return typeof definition === "string" ? definition : definition.value;
}

function shortcutAfter(definition: VisualTexInlineShortcutDefinition) {
  return typeof definition === "string" ? "" : definition.after ?? "";
}

function previewLatex(definition: VisualTexInlineShortcutDefinition) {
  return shortcutLatex(definition).replaceAll("#?", "\\square");
}

function readActiveInlineShortcuts(): VisualTexInlineShortcutDefinitions {
  try {
    const field = document.querySelector("math-field") as
      | (HTMLElement & {
          inlineShortcuts?: Readonly<VisualTexInlineShortcutDefinitions>;
        })
      | null;
    const active = field?.inlineShortcuts;
    if (active && Object.keys(active).length > 0) return { ...active };
  } catch {
    // Fall back to the explicit VisualTeX table while the mathfield mounts.
  }
  return { ...visualTexAutoEscapeInlineShortcuts };
}

interface InputBehaviorPopoverLayout {
  left: number;
  top: number;
  width: number;
  maxHeight: number;
  compact: boolean;
}

function ShortcutOutput({ definition }: { definition: VisualTexInlineShortcutDefinition }) {
  const latex = previewLatex(definition);
  const markup = useMemo(() => {
    try {
      return convertVisualTexLatexToMarkup(latex, { defaultMode: "math" });
    } catch {
      return "";
    }
  }, [latex]);

  if (!markup) return <code>{shortcutLatex(definition)}</code>;
  return (
    <span
      className="auto-escape-map-output-formula"
      dangerouslySetInnerHTML={{ __html: markup }}
    />
  );
}

export function InputBehaviorMenu() {
  const [open, setOpen] = useState(false);
  const [page, setPage] = useState<"settings" | "mappings">("settings");
  const [mappingQuery, setMappingQuery] = useState("");
  const [activeShortcutDefinitions, setActiveShortcutDefinitions] =
    useState<VisualTexInlineShortcutDefinitions>(() => ({
      ...visualTexAutoEscapeInlineShortcuts,
    }));
  const rootRef = useRef<HTMLDivElement>(null);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const popoverRef = useRef<HTMLDivElement>(null);
  const [popoverLayout, setPopoverLayout] =
    useState<InputBehaviorPopoverLayout | null>(null);
  const language = useEditorStore((state) => state.language);
  const inputBehavior = useEditorStore((state) => state.inputBehavior);
  const setInputBehavior = useEditorStore((state) => state.setInputBehavior);
  const isEn = language === "en";

  const filteredShortcutGroups = useMemo(() => {
    const query = mappingQuery.trim().toLocaleLowerCase();
    const seen = new Set<string>();
    const groups = visualTexAutoEscapeShortcutGroups.map((group) => {
      const entries = Object.keys(group.shortcuts)
        .filter((shortcut) => shortcut in activeShortcutDefinitions)
        .map(
          (shortcut) =>
            [shortcut, activeShortcutDefinitions[shortcut]] as const,
        );
      entries.forEach(([shortcut]) => seen.add(shortcut));
      return { ...group, entries };
    });
    const mathLiveEntries = Object.entries(activeShortcutDefinitions).filter(
      ([shortcut]) => !seen.has(shortcut),
    );
    if (mathLiveEntries.length > 0) {
      groups.push({
        id: "mathlive",
        titleZh: "MathLive 内置",
        titleEn: "MathLive built-ins",
        shortcuts: {},
        entries: mathLiveEntries,
      });
    }
    return groups
      .map((group) => ({
        ...group,
        entries: group.entries.filter(([shortcut, definition]) => {
          if (!query) return true;
          return (
            shortcut.toLocaleLowerCase().includes(query) ||
            shortcutLatex(definition).toLocaleLowerCase().includes(query)
          );
        }),
      }))
      .filter((group) => group.entries.length > 0);
  }, [activeShortcutDefinitions, mappingQuery]);

  const mappingCount = useMemo(
    () => filteredShortcutGroups.reduce((sum, group) => sum + group.entries.length, 0),
    [filteredShortcutGroups],
  );

  useEffect(() => {
    if (!open) return;
    const handlePointerDown = (event: PointerEvent) => {
      const target = event.target as Node;
      if (
        !rootRef.current?.contains(target) &&
        !popoverRef.current?.contains(target)
      ) {
        setOpen(false);
      }
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      if (page === "mappings") {
        setPage("settings");
        setMappingQuery("");
      } else {
        setOpen(false);
      }
    };
    document.addEventListener("pointerdown", handlePointerDown, true);
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("pointerdown", handlePointerDown, true);
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [open, page]);

  useEffect(() => {
    if (open) return;
    setPage("settings");
    setMappingQuery("");
    setPopoverLayout(null);
  }, [open]);

  useLayoutEffect(() => {
    if (!open) return;
    const trigger = triggerRef.current;
    if (!trigger) return;

    let frame = 0;
    const updateLayout = () => {
      window.cancelAnimationFrame(frame);
      frame = window.requestAnimationFrame(() => {
        const triggerRect = trigger.getBoundingClientRect();
        const workspace = trigger.closest<HTMLElement>(".workspace");
        const visibleEditor = workspace?.classList.contains("is-classic-layout")
          ? workspace.querySelector<HTMLElement>(".classic-editor-pane-body")
          : workspace?.querySelector<HTMLElement>(".formula-workspace.editor-pane");
        const editorRect = visibleEditor?.getBoundingClientRect();
        const viewportMargin = 8;
        const popoverGap = 6;
        const viewportRight = Math.max(viewportMargin, window.innerWidth - viewportMargin);
        const viewportBottom = Math.max(viewportMargin, window.innerHeight - viewportMargin);
        const editorIsUsable = Boolean(
          editorRect && editorRect.width >= 140 && editorRect.height >= 100,
        );
        const leftBound = editorIsUsable
          ? Math.max(viewportMargin, editorRect!.left + viewportMargin)
          : viewportMargin;
        const rightBound = editorIsUsable
          ? Math.min(viewportRight, editorRect!.right - viewportMargin)
          : viewportRight;
        const availableWidth = Math.max(120, rightBound - leftBound);
        const preferredWidth = page === "mappings" ? 660 : 420;
        const width = Math.min(preferredWidth, availableWidth);
        const left = Math.min(
          Math.max(triggerRect.right - width, leftBound),
          Math.max(leftBound, rightBound - width),
        );
        const minimumTop = editorIsUsable
          ? Math.max(viewportMargin, editorRect!.top + viewportMargin)
          : viewportMargin;
        const top = Math.min(
          Math.max(triggerRect.bottom + popoverGap, minimumTop),
          Math.max(viewportMargin, viewportBottom - 96),
        );
        const preferredMaxHeight = page === "mappings" ? 620 : 560;
        const maxHeight = Math.max(
          96,
          Math.min(preferredMaxHeight, viewportBottom - top),
        );
        const next = {
          left,
          top,
          width,
          maxHeight,
          compact: width < 360 || maxHeight < 420,
        };
        setPopoverLayout((current) =>
          current &&
          Math.abs(current.left - next.left) < 0.5 &&
          Math.abs(current.top - next.top) < 0.5 &&
          Math.abs(current.width - next.width) < 0.5 &&
          Math.abs(current.maxHeight - next.maxHeight) < 0.5 &&
          current.compact === next.compact
            ? current
            : next,
        );
      });
    };

    const resizeObserver = new ResizeObserver(updateLayout);
    resizeObserver.observe(trigger);
    const workspace = trigger.closest<HTMLElement>(".workspace");
    if (workspace) resizeObserver.observe(workspace);
    window.addEventListener("resize", updateLayout);
    window.addEventListener("scroll", updateLayout, true);
    updateLayout();
    return () => {
      window.cancelAnimationFrame(frame);
      resizeObserver.disconnect();
      window.removeEventListener("resize", updateLayout);
      window.removeEventListener("scroll", updateLayout, true);
    };
  }, [open, page]);

  const popoverStyle = popoverLayout
    ? ({
        left: `${popoverLayout.left}px`,
        top: `${popoverLayout.top}px`,
        width: `${popoverLayout.width}px`,
        maxWidth: `${popoverLayout.width}px`,
        maxHeight: `${popoverLayout.maxHeight}px`,
      } as CSSProperties)
    : undefined;

  return (
    <div ref={rootRef} className="input-behavior-menu">
      <button
        ref={triggerRef}
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

      {open &&
        popoverLayout &&
        createPortal(
        <div
          ref={popoverRef}
          className={
            `input-behavior-popover${page === "mappings" ? " is-mapping-view" : ""}` +
            (popoverLayout.compact ? " is-compact" : "")
          }
          style={popoverStyle}
          role="dialog"
          aria-label={
            page === "mappings"
              ? isEn
                ? "Automatic conversion mappings"
                : "自动转义映射"
              : isEn
                ? "Input behavior settings"
                : "操作逻辑设置"
          }
        >
          {page === "mappings" ? (
            <>
              <div className="auto-escape-map-toolbar">
                <button
                  type="button"
                  className="auto-escape-map-back"
                  aria-label={isEn ? "Back to input behavior" : "返回操作逻辑"}
                  onClick={() => {
                    setPage("settings");
                    setMappingQuery("");
                  }}
                >
                  <ArrowLeft size={15} />
                </button>
                <div>
                  <strong>{isEn ? "Conversion mappings" : "自动转义映射"}</strong>
                  <span>{mappingCount}</span>
                </div>
              </div>

              <label className="auto-escape-map-search">
                <Search size={14} aria-hidden="true" />
                <input
                  value={mappingQuery}
                  onChange={(event) => setMappingQuery(event.target.value)}
                  placeholder={isEn ? "Search input or LaTeX" : "搜索输入或 LaTeX"}
                  aria-label={isEn ? "Search conversion mappings" : "搜索自动转义映射"}
                  autoFocus
                />
              </label>

              <div className="auto-escape-map-groups">
                {filteredShortcutGroups.map((group) => (
                  <section className="auto-escape-map-group" key={group.id}>
                    <div className="auto-escape-map-group-heading">
                      <strong>{isEn ? group.titleEn : group.titleZh}</strong>
                      {group.entries.some(([, definition]) => shortcutAfter(definition)) ? (
                        <span
                          title={
                            isEn
                              ? "These entries apply only after the allowed preceding structures defined in code"
                              : "这些条目只在代码规定的前置结构后触发"
                          }
                        >
                          {isEn ? "context" : "有前置条件"}
                        </span>
                      ) : null}
                    </div>
                    <div className="auto-escape-map-grid">
                      {group.entries.map(([shortcut, definition]) => (
                        <div
                          className="auto-escape-map-row"
                          key={shortcut}
                          data-auto-escape-shortcut={shortcut}
                          data-auto-escape-output={shortcutLatex(definition)}
                          data-auto-escape-after={shortcutAfter(definition)}
                          title={`${shortcut} → ${shortcutLatex(definition)}${
                            shortcutAfter(definition)
                              ? ` · after: ${shortcutAfter(definition)}`
                              : ""
                          }`}
                        >
                          <span className="auto-escape-map-input-wrap">
                            <code className="auto-escape-map-input">{shortcut}</code>
                            {shortcutAfter(definition) ? (
                              <span
                                className="auto-escape-condition-mark"
                                aria-label={isEn ? "Has context condition" : "有前置条件"}
                              />
                            ) : null}
                          </span>
                          <ArrowRight size={13} aria-hidden="true" />
                          <span className="auto-escape-map-output">
                            <ShortcutOutput definition={definition} />
                          </span>
                        </div>
                      ))}
                    </div>
                  </section>
                ))}
                {filteredShortcutGroups.length === 0 ? (
                  <div className="auto-escape-map-empty">
                    {isEn ? "No matching mapping" : "没有匹配的映射"}
                  </div>
                ) : null}
              </div>
            </>
          ) : (
            <>
              <div className="input-behavior-heading">
                <strong>{isEn ? "Automatic conversion" : "输入自动转义"}</strong>
              </div>

              <div className="input-behavior-options">
                {AUTO_ESCAPE_OPTIONS.map((option) => (
                  <div className="input-behavior-option has-secondary-action" key={option.key}>
                    <span>
                      <strong>{isEn ? option.titleEn : option.titleZh}</strong>
                      <button
                        type="button"
                        className="input-behavior-map-button"
                        data-open-auto-escape-map
                        onClick={() => {
                          setActiveShortcutDefinitions(readActiveInlineShortcuts());
                          setPage("mappings");
                        }}
                      >
                        {isEn ? "View mappings" : "查看映射"}
                        <ArrowRight size={13} aria-hidden="true" />
                      </button>
                    </span>
                    <label
                      className="input-behavior-toggle"
                      aria-label={isEn ? option.titleEn : option.titleZh}
                    >
                      <input
                        type="checkbox"
                        checked={inputBehavior[option.key]}
                        onChange={(event) =>
                          setInputBehavior(option.key, event.target.checked)
                        }
                      />
                      <span className="input-behavior-switch" aria-hidden="true" />
                    </label>
                  </div>
                ))}
              </div>

              <div className="input-behavior-heading input-behavior-section-heading">
                <strong>{isEn ? "Caret auto-exit" : "光标自动跳出"}</strong>
              </div>

              <div className="input-behavior-options">
                {CARET_BEHAVIOR_OPTIONS.map((option) => (
                  <label className="input-behavior-option" key={option.key}>
                    <span>
                      <strong>{isEn ? option.titleEn : option.titleZh}</strong>
                    </span>
                    <input
                      type="checkbox"
                      checked={inputBehavior[option.key]}
                      onChange={(event) => setInputBehavior(option.key, event.target.checked)}
                    />
                    <span className="input-behavior-switch" aria-hidden="true" />
                  </label>
                ))}
              </div>

              <div className="input-behavior-heading input-behavior-section-heading">
                <strong>{isEn ? "Command suggestion panels" : "命令候选框"}</strong>
              </div>

              <div className="input-behavior-options">
                {COMMAND_SUGGESTION_OPTIONS.map((option) => (
                  <label className="input-behavior-option" key={option.key}>
                    <span>
                      <strong>{isEn ? option.titleEn : option.titleZh}</strong>
                    </span>
                    <input
                      type="checkbox"
                      checked={inputBehavior[option.key]}
                      onChange={(event) => setInputBehavior(option.key, event.target.checked)}
                    />
                    <span className="input-behavior-switch" aria-hidden="true" />
                  </label>
                ))}
              </div>
            </>
          )}
        </div>,
        document.body,
      )}
    </div>
  );
}
