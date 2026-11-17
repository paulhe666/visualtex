import { useEffect, useMemo, useState } from "react";
import { createPortal } from "react-dom";
import { Check, Plus, Search, X } from "lucide-react";
import { MathPreview } from "../components/MathPreview";
import type { LatexCommand } from "../types/command";
import { categoryLabels, categoryLabelsEn, commandRegistry } from "../autocomplete/commandRegistry";
import { customSymbolCommands } from "../autocomplete/runtimeCommandRegistry";
import { BUILT_IN_GROUP_IDS, defaultGroupItems, hiddenToolbarCommandIds, type ToolbarItem } from "./toolbarLayout";

interface Props {
  open: boolean;
  language: string;
  /** the group the picked things go to */
  groupName: string;
  /** the group holds this thing */
  has: (item: ToolbarItem) => boolean;
  onToggle: (item: ToolbarItem) => void;
  /** the formula being edited (offered as one more card) */
  currentLatex: string;
  onClose: () => void;
}

type Source = { kind: "category"; id: string } | { kind: "symbols" };

interface Card {
  key: string;
  item: ToolbarItem;
  latex: string;
  label: string;
}

const registryById = new Map(commandRegistry.map((command) => [command.id, command]));

function commandCard(command: LatexCommand, isEn: boolean, preview?: string): Card {
  return { key: `command:${command.id}`, item: { kind: "command", id: command.id }, latex: preview ?? command.previewLatex, label: isEn ? command.labelEn : command.labelZh };
}

/**
 * What a group of the toolbar can be given: the toolbar's commands by category and the custom symbols, as a dense field of
 * cards that show the character and nothing else (the name is the tooltip). The formula tiles' library is the tiles' own
 * dialog's: it is not offered here.
 */
export function ToolbarItemPickerDialog({ open, language, groupName, has, onToggle, currentLatex, onClose }: Props) {
  const isEn = language === "en";
  const [source, setSource] = useState<Source>({ kind: "category", id: "common" });
  const [query, setQuery] = useState("");

  useEffect(() => {
    if (!open) return;
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      event.stopPropagation();
      onClose();
    };
    window.addEventListener("keydown", closeOnEscape, true);
    return () => window.removeEventListener("keydown", closeOnEscape, true);
  }, [onClose, open]);

  const symbols = useMemo(() => (open ? customSymbolCommands() : []), [open]);
  const needle = query.trim().toLowerCase();

  const cards = useMemo<Card[]>(() => {
    const matrixCard: Card = { key: "matrix", item: { kind: "matrix" }, latex: "\\begin{bmatrix}a&b\\\\c&d\\end{bmatrix}", label: isEn ? "Matrix builder" : "矩阵生成器" };
    const categoryCards = (id: string): Card[] =>
      defaultGroupItems(id as (typeof BUILT_IN_GROUP_IDS)[number]).flatMap((item): Card[] => {
        if (item.kind === "matrix") return [matrixCard];
        if (item.kind !== "command") return [];
        const command = registryById.get(item.id);
        return command ? [commandCard(command, isEn)] : [];
      });
    if (needle) {
      const matches = (text: string) => text.toLowerCase().includes(needle);
      const commandHits = commandRegistry
        .filter((command) => !hiddenToolbarCommandIds.has(command.id))
        .filter((command) => matches(command.labelZh) || matches(command.labelEn) || matches(command.command) || command.aliases.some(matches) || command.keywords.some(matches))
        .slice(0, 240)
        .map((command) => commandCard(command, isEn));
      const symbolHits = symbols.filter((command) => matches(command.labelZh) || matches(command.command)).map((command) => commandCard(command, isEn));
      return [...(matches(matrixCard.label) || matches("matrix") || matches("矩阵") ? [matrixCard] : []), ...commandHits, ...symbolHits];
    }
    if (source.kind === "symbols") return symbols.map((command) => commandCard(command, isEn));
    return categoryCards(source.id);
  }, [isEn, needle, source, symbols]);

  if (!open) return null;

  const current: Card | null = currentLatex.trim()
    ? { key: "current", item: { kind: "formula", id: `formula-${Date.now()}`, latex: currentLatex.trim() }, latex: currentLatex.trim(), label: isEn ? "Current formula" : "当前公式" }
    : null;
  const isActive = (candidate: Source) =>
    !needle && candidate.kind === source.kind && ("id" in candidate ? "id" in source && candidate.id === source.id : true);

  const renderCard = (card: Card) => {
    const added = has(card.item);
    return (
      <button
        key={card.key}
        type="button"
        className={"formula-tile-library-item toolbar-picker-item" + (card.latex.length > 72 ? " is-wide" : "") + (added ? " is-added" : "")}
        aria-pressed={added}
        aria-label={card.label}
        data-toolbar-picker-item={card.key}
        title={card.label}
        onClick={() => onToggle(card.item)}
      >
        <MathPreview latex={card.latex} className="formula-tile-library-preview" fit maximumFitScale={1} />
        {added && (
          <span className="formula-tile-library-item-state" aria-hidden="true">
            <Check size={12} />
          </span>
        )}
      </button>
    );
  };

  return createPortal(
    <div className="modal-backdrop formula-tile-library-backdrop" role="presentation" onMouseDown={onClose}>
      <section className="formula-tile-library-dialog toolbar-item-picker" role="dialog" aria-modal="true" aria-labelledby="toolbar-item-picker-title" data-toolbar-item-picker onMouseDown={(event) => event.stopPropagation()}>
        <header className="dialog-header">
          <div>
            <h2 id="toolbar-item-picker-title">{isEn ? `Add to “${groupName}”` : `添加到「${groupName}」`}</h2>
          </div>
          <label className="formula-tile-library-search">
            <Search size={14} aria-hidden="true" />
            <input
              type="search"
              autoFocus
              value={query}
              placeholder={isEn ? "Search" : "搜索"}
              aria-label={isEn ? "Search" : "搜索"}
              onChange={(event) => setQuery(event.currentTarget.value)}
            />
          </label>
          <button type="button" className="icon-button" onClick={onClose} aria-label={isEn ? "Close" : "关闭"}>
            <X size={18} />
          </button>
        </header>
        <div className="formula-tile-library-body">
          <nav className="formula-tile-library-nav" aria-label={isEn ? "Sources" : "来源"}>
            <div className="formula-tile-library-groups">
              {BUILT_IN_GROUP_IDS.map((id) => (
                <button
                  key={id}
                  type="button"
                  className={isActive({ kind: "category", id }) ? "is-active" : ""}
                  data-toolbar-picker-source={id}
                  onClick={() => {
                    setQuery("");
                    setSource({ kind: "category", id });
                  }}
                >
                  <span className="formula-tile-library-group-name">{(isEn ? categoryLabelsEn : categoryLabels)[id]}</span>
                </button>
              ))}
              {symbols.length > 0 && (
                <button
                  type="button"
                  className={isActive({ kind: "symbols" }) ? "is-active" : ""}
                  data-toolbar-picker-source="symbols"
                  onClick={() => {
                    setQuery("");
                    setSource({ kind: "symbols" });
                  }}
                >
                  <span className="formula-tile-library-group-name">{isEn ? "Custom symbols" : "自定义字符"}</span>
                </button>
              )}
            </div>
          </nav>
          <div className="formula-tile-library-content">
            {cards.length === 0 ? <p className="formula-tile-library-empty">{isEn ? "Nothing matches." : "没有匹配的内容。"}</p> : <div className="formula-tile-library-grid">{cards.map(renderCard)}</div>}
          </div>
        </div>
        <footer className="formula-tile-library-footer">
          {current ? (
            <button type="button" className="secondary-button" data-toolbar-picker-current onClick={() => onToggle(current.item)} disabled={has(current.item)}>
              <Plus size={14} />
              {isEn ? "Add current formula" : "添加当前公式"}
            </button>
          ) : (
            <span />
          )}
          <button type="button" className="secondary-button" onClick={onClose}>
            {isEn ? "Done" : "完成"}
          </button>
        </footer>
      </section>
    </div>,
    document.body,
  );
}
