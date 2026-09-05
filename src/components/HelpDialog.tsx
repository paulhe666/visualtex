import { Fragment, useEffect, useMemo, type ReactNode } from "react";
import { BookOpenText, X } from "lucide-react";
import manualMarkdown from "../../docs/help/VisualTeX-Help.md?raw";

interface HelpDialogProps {
  open: boolean;
  language: "cn" | "en";
  onClose: () => void;
}

type HelpBlock =
  | { type: "heading"; level: 1 | 2 | 3; text: string; id?: string }
  | { type: "paragraph"; text: string }
  | { type: "unordered"; items: string[] }
  | { type: "ordered"; items: string[] }
  | { type: "code"; language: string; text: string }
  | { type: "table"; headers: string[]; rows: string[][] }
  | { type: "rule" };

function tableCells(line: string) {
  return line
    .trim()
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split("|")
    .map((cell) => cell.trim());
}

function isTableDivider(line: string) {
  const cells = tableCells(line);
  return cells.length > 0 && cells.every((cell) => /^:?-{3,}:?$/.test(cell));
}

function parseManual(markdown: string): HelpBlock[] {
  const lines = markdown.replace(/\r\n?/g, "\n").split("\n");
  const blocks: HelpBlock[] = [];
  let index = 0;
  let sectionIndex = 0;

  while (index < lines.length) {
    const raw = lines[index];
    const line = raw.trim();
    if (!line) {
      index += 1;
      continue;
    }

    if (line.startsWith("```")) {
      const language = line.slice(3).trim();
      const code: string[] = [];
      index += 1;
      while (index < lines.length && !lines[index].trim().startsWith("```")) {
        code.push(lines[index]);
        index += 1;
      }
      if (index < lines.length) index += 1;
      blocks.push({ type: "code", language, text: code.join("\n") });
      continue;
    }

    const heading = /^(#{1,3})\s+(.+)$/.exec(line);
    if (heading) {
      const level = heading[1].length as 1 | 2 | 3;
      const text = heading[2].trim();
      const id = level === 2 ? `help-section-${++sectionIndex}` : undefined;
      blocks.push({ type: "heading", level, text, id });
      index += 1;
      continue;
    }

    if (/^-{3,}$/.test(line)) {
      blocks.push({ type: "rule" });
      index += 1;
      continue;
    }

    if (
      line.includes("|") &&
      index + 1 < lines.length &&
      isTableDivider(lines[index + 1])
    ) {
      const headers = tableCells(raw);
      const rows: string[][] = [];
      index += 2;
      while (index < lines.length && lines[index].trim().includes("|")) {
        rows.push(tableCells(lines[index]));
        index += 1;
      }
      blocks.push({ type: "table", headers, rows });
      continue;
    }

    if (/^-\s+/.test(line)) {
      const items: string[] = [];
      while (index < lines.length && /^-\s+/.test(lines[index].trim())) {
        items.push(lines[index].trim().replace(/^-\s+/, ""));
        index += 1;
      }
      blocks.push({ type: "unordered", items });
      continue;
    }

    if (/^\d+\.\s+/.test(line)) {
      const items: string[] = [];
      while (index < lines.length && /^\d+\.\s+/.test(lines[index].trim())) {
        items.push(lines[index].trim().replace(/^\d+\.\s+/, ""));
        index += 1;
      }
      blocks.push({ type: "ordered", items });
      continue;
    }

    const paragraph: string[] = [line];
    index += 1;
    while (index < lines.length) {
      const next = lines[index].trim();
      if (
        !next ||
        next.startsWith("```") ||
        /^(#{1,3})\s+/.test(next) ||
        /^-{3,}$/.test(next) ||
        /^-\s+/.test(next) ||
        /^\d+\.\s+/.test(next) ||
        (next.includes("|") &&
          index + 1 < lines.length &&
          isTableDivider(lines[index + 1]))
      ) {
        break;
      }
      paragraph.push(next);
      index += 1;
    }
    blocks.push({ type: "paragraph", text: paragraph.join(" ") });
  }

  return blocks;
}

function renderInline(text: string): ReactNode {
  return text
    .split(/(\*\*[^*]+\*\*|`[^`]+`)/g)
    .filter(Boolean)
    .map((part, index) => {
      if (part.startsWith("**") && part.endsWith("**")) {
        return <strong key={index}>{part.slice(2, -2)}</strong>;
      }
      if (part.startsWith("`") && part.endsWith("`")) {
        return <code key={index}>{part.slice(1, -1)}</code>;
      }
      return <Fragment key={index}>{part}</Fragment>;
    });
}

export function HelpDialog({ open, language, onClose }: HelpDialogProps) {
  const blocks = useMemo(() => parseManual(manualMarkdown), []);
  const sections = useMemo(
    () =>
      blocks.filter(
        (block): block is Extract<HelpBlock, { type: "heading" }> =>
          block.type === "heading" && block.level === 2 && Boolean(block.id),
      ),
    [blocks],
  );
  const isEn = language === "en";

  useEffect(() => {
    if (!open) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") onClose();
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [open, onClose]);

  if (!open) return null;

  return (
    <div
      className="modal-backdrop help-dialog-backdrop"
      role="presentation"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onClose();
      }}
    >
      <section
        className="dialog help-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="help-dialog-title"
      >
        <header className="help-dialog-header">
          <div>
            <strong id="help-dialog-title">
              <BookOpenText size={18} />
              {isEn ? "Help Manual" : "帮助手册"}
            </strong>
            <span>VisualTeX Web</span>
          </div>
          <button
            type="button"
            className="icon-button compact"
            onClick={onClose}
            aria-label={isEn ? "Close help manual" : "关闭帮助手册"}
          >
            <X size={17} />
          </button>
        </header>

        <div className="help-dialog-body">
          <nav className="help-dialog-toc" aria-label={isEn ? "Manual sections" : "手册目录"}>
            {sections.map((section) => (
              <button
                type="button"
                key={section.id}
                onClick={() =>
                  document.getElementById(section.id!)?.scrollIntoView({
                    behavior: "smooth",
                    block: "start",
                  })
                }
              >
                {section.text}
              </button>
            ))}
          </nav>

          <article className="help-dialog-content">
            {blocks.map((block, index) => {
              if (block.type === "heading") {
                const Tag = `h${block.level}` as "h1" | "h2" | "h3";
                return (
                  <Tag key={index} id={block.id}>
                    {renderInline(block.text)}
                  </Tag>
                );
              }
              if (block.type === "paragraph") {
                return <p key={index}>{renderInline(block.text)}</p>;
              }
              if (block.type === "unordered") {
                return (
                  <ul key={index}>
                    {block.items.map((item, itemIndex) => (
                      <li key={itemIndex}>{renderInline(item)}</li>
                    ))}
                  </ul>
                );
              }
              if (block.type === "ordered") {
                return (
                  <ol key={index}>
                    {block.items.map((item, itemIndex) => (
                      <li key={itemIndex}>{renderInline(item)}</li>
                    ))}
                  </ol>
                );
              }
              if (block.type === "code") {
                return (
                  <pre key={index} data-language={block.language || undefined}>
                    <code>{block.text}</code>
                  </pre>
                );
              }
              if (block.type === "table") {
                return (
                  <div className="help-table-scroll" key={index}>
                    <table>
                      <thead>
                        <tr>
                          {block.headers.map((header, cellIndex) => (
                            <th key={cellIndex}>{renderInline(header)}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {block.rows.map((row, rowIndex) => (
                          <tr key={rowIndex}>
                            {block.headers.map((_, cellIndex) => (
                              <td key={cellIndex}>{renderInline(row[cellIndex] ?? "")}</td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                );
              }
              return <hr key={index} />;
            })}
          </article>
        </div>
      </section>
    </div>
  );
}
