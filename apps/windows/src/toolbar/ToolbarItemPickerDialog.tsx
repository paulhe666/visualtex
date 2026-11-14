import { useEffect, useMemo, useState } from "react";
import { createPortal } from "react-dom";
import { Check, Plus, Search, X } from "lucide-react";
import { MathPreview } from "../components/MathPreview";
import type { LatexCommand } from "../../../shared/visualtex-app/src/types/command";
import { categoryLabels, categoryLabelsEn, commandRegistry } from "../../../shared/visualtex-app/src/autocomplete/commandRegistry";
import { customSymbolCommands } from "../../../shared/visualtex-app/src/autocomplete/runtimeCommandRegistry";
import { formulaTileLevelLabels, formulaTileLibrary, formulaTileSubjects } from "../../../shared/visualtex-app/src/toolbar/formulaTileLibrary";
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

type Source = { kind: "category"; id: string } | { kind: "symbols" } | { kind: "library"; id: string };

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

/** What a group of the toolbar can be given: the toolbar's commands by category, the custom symbols, the formula library. */
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

  const sections = useMemo<Array<{ id: string; title: string; cards: Card[] }>>(() => {
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
        .slice(0, 160)
        .map((command) => commandCard(command, isEn));
      const symbolHits = symbols.filter((command) => matches(command.labelZh) || matches(command.command)).map((command) => commandCard(command, isEn));
      const libraryHits = formulaTileLibrary
        .flatMap((group) => group.tiles)
        .filter((tile) => matches(tile.labelZh) || matches(tile.labelEn) || matches(tile.latex))
        .slice(0, 160)
        .map((tile): Card => ({ key: `library:${tile.id}`, item: { kind: "formula", id: `library-${tile.id}`, latex: tile.latex }, latex: tile.latex, label: isEn ? tile.labelEn : tile.labelZh }));
      return [
        { id: "commands", title: isEn ? "Toolbar commands" : "工具栏命令", cards: [...(matches(matrixCard.label) || matches("matrix") || matches("矩阵") ? [matrixCard] : []), ...commandHits] },
        { id: "symbols", title: isEn ? "Custom symbols" : "自定义字符", cards: symbolHits },
        { id: "library", title: isEn ? "Formula library" : "公式库", cards: libraryHits },
      ].filter((section) => section.cards.length > 0);
    }
    if (source.kind === "symbols") return [{ id: "symbols", title: isEn ? "Custom symbols" : "自定义字符", cards: symbols.map((command) => commandCard(command, isEn)) }];
    if (source.kind === "library") {
      const group = formulaTileLibrary.find((candidate) => candidate.id === source.id) ?? formulaTileLibrary[0];
      return [
        {
          id: group.id,
          title: isEn ? group.labelEn : group.labelZh,
          cards: group.tiles.map((tile) => ({ key: `library:${tile.id}`, item: { kind: "formula", id: `library-${tile.id}`, latex: tile.latex }, latex: tile.latex, label: isEn ? tile.labelEn : tile.labelZh })),
        },
      ];
    }
    return [{ id: source.id, title: (isEn ? categoryLabelsEn : categoryLabels)[source.id] ?? source.id, cards: categoryCards(source.id) }];
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
        data-toolbar-picker-item={card.key}
        title={card.label}
        onClick={() => onToggle(card.item)}
      >
        <MathPreview latex={card.latex} className="formula-tile-library-preview" fit maximumFitScale={1} />
        <span className="formula-tile-library-item-label">{card.label}</span>
        <span className="formula-tile-library-item-state" aria-hidden="true">
          {added ? <Check size={12} /> : <Plus size={12} />}
        </span>
      </button>
    );
  };

  return createPortal(
    <div className="modal-backdrop formula-tile-library-backdrop" role="presentation" onMouseDown={onClose}>
      <section className="formula-tile-library-dialog toolbar-item-picker" role="dialog" aria-modal="true" aria-labelledby="toolbar-item-picker-title" data-toolbar-item-picker onMouseDown={(event) => event.stopPropagation()}>
        <header className="dialog-header">
          <div>
            <span className="eyebrow">{isEn ? "TOOLBAR GROUP" : "工具栏分组"}</span>
            <h2 id="toolbar-item-picker-title">{isEn ? `Add to “${groupName}”` : `添加到「${groupName}」`}</h2>
          </div>
          <label className="formula-tile-library-search">
            <Search size={14} aria-hidden="true" />
            <input
              type="search"
              autoFocus
              value={query}
              placeholder={isEn ? "Search commands and formulas" : "搜索命令和公式"}
              aria-label={isEn ? "Search commands and formulas" : "搜索命令和公式"}
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
              <span className="toolbar-picker-nav-heading">{isEn ? "Toolbar commands" : "工具栏命令"}</span>
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
                  <span className="formula-tile-library-group-meta">{symbols.length}</span>
                </button>
              )}
              {formulaTileSubjects.map((subject) => (
                <div key={subject.id} className="toolbar-picker-nav-block">
                  <span className="toolbar-picker-nav-heading">{isEn ? subject.labelEn : subject.labelZh}</span>
                  {formulaTileLibrary
                    .filter((group) => group.subject === subject.id)
                    .map((group) => (
                      <button
                        key={group.id}
                        type="button"
                        className={isActive({ kind: "library", id: group.id }) ? "is-active" : ""}
                        data-toolbar-picker-source={group.id}
                        onClick={() => {
                          setQuery("");
                          setSource({ kind: "library", id: group.id });
                        }}
                      >
                        <span className="formula-tile-library-group-name">{isEn ? group.labelEn : group.labelZh}</span>
                        <span className="formula-tile-library-group-meta">
                          {isEn ? formulaTileLevelLabels[group.level].labelEn : formulaTileLevelLabels[group.level].labelZh}
                        </span>
                      </button>
                    ))}
                </div>
              ))}
            </div>
          </nav>
          <div className="formula-tile-library-content">
            {sections.length === 0 && <p className="formula-tile-library-empty">{isEn ? "Nothing matches." : "没有匹配的内容。"}</p>}
            {sections.map((section) => (
              <section key={section.id} className="formula-tile-library-section">
                <h3>{section.title}</h3>
                <div className="formula-tile-library-grid">{section.cards.map(renderCard)}</div>
              </section>
            ))}
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
