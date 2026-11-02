// LaTeX source read into the document (a paste, an opened .tex file): it is turned into the Markdown the document
// already reads (markdown.ts) — sections into headings, \textbf / \emph / \underline into marks, itemize /
// enumerate into lists, quote into a quote, tabular into a table, lstlisting / verbatim into code blocks, math as
// it is ($…$, \(…\), \[…\], equation / align / gather as numbered display formulas, \eqref as references).
// Commands without a counterpart (\maketitle, \usepackage, \centering…) are dropped; their text stays.

/** Whether text is LaTeX document source (not just a formula). */
export function looksLikeLatexSource(text: string): boolean {
  return (
    /\\documentclass|\\begin\{document\}/.test(text) ||
    /\\(?:chapter|section|subsection|subsubsection|paragraph)\*?\s*\{/.test(text) ||
    /\\begin\{(?:itemize|enumerate|description|quote|quotation|tabular|table|figure|lstlisting|verbatim|center|abstract)\}/.test(text) ||
    /\\(?:textbf|textit|emph|underline|texttt)\s*\{/.test(text)
  );
}

/** The {…} group starting at `open` (a "{"): its content and the index after its "}". */
function group(text: string, open: number): { content: string; end: number } | null {
  if (text[open] !== "{") return null;
  let depth = 0;
  for (let i = open; i < text.length; i++) {
    const c = text[i];
    if (c === "\\") {
      i++;
      continue;
    }
    if (c === "{") depth++;
    else if (c === "}" && --depth === 0) return { content: text.slice(open + 1, i), end: i + 1 };
  }
  return null;
}

/** Skip spaces (not blank lines) from i. */
function skipSpaces(text: string, i: number): number {
  while (i < text.length && (text[i] === " " || text[i] === "\t")) i++;
  return i;
}

/** An optional [...] argument at i: the index after it. */
function skipOptional(text: string, i: number): number {
  const j = skipSpaces(text, i);
  if (text[j] !== "[") return i;
  const close = text.indexOf("]", j);
  return close < 0 ? i : close + 1;
}

const DISPLAY_ENVIRONMENTS = "equation|align|gather|multline|eqnarray|flalign|alignat";

/** Math kept aside while the text around it is converted (placeholders \u0000n\u0000). */
function protectMath(text: string, kept: string[]): string {
  const keep = (markdown: string) => {
    kept.push(markdown);
    return `\u0000${kept.length - 1}\u0000`;
  };
  // Display environments: numbered unless starred; align and gather become aligned and gathered inside $$.
  text = text.replace(new RegExp(`\\\\begin\\{(${DISPLAY_ENVIRONMENTS})(\\*?)\\}([\\s\\S]*?)\\\\end\\{\\1\\2\\}`, "g"), (_m, env: string, star: string, body: string) => {
    const inner = body.trim();
    if (star) {
      const name = env === "align" || env === "flalign" || env === "alignat" || env === "eqnarray" ? "aligned" : env === "gather" ? "gathered" : null;
      return keep(`\n\n$$\n${name ? `\\begin{${name}}${inner}\\end{${name}}` : inner}\n$$\n\n`);
    }
    const kind = env === "equation" || env === "multline" ? "equation" : env === "gather" ? "gather" : "align";
    return keep(`\n\n$$\n\\begin{${kind}}${inner}\\end{${kind}}\n$$\n\n`);
  });
  text = text.replace(/\$\$([\s\S]+?)\$\$/g, (_m, body: string) => keep(`\n\n$$\n${body.trim()}\n$$\n\n`));
  text = text.replace(/\\\[([\s\S]+?)\\\]/g, (_m, body: string) => keep(`\n\n$$\n${body.trim()}\n$$\n\n`));
  text = text.replace(/\\\(([\s\S]+?)\\\)/g, (_m, body: string) => keep(`$${body.trim()}$`));
  text = text.replace(/(?<!\\)\$([^$]+?)(?<!\\)\$/g, (_m, body: string) => keep(`$${body.trim()}$`));
  return text;
}

/** Code environments, kept verbatim as fences. */
function protectCode(text: string, kept: string[]): string {
  return text.replace(/\\begin\{(lstlisting|verbatim|minted)\}(\[[^\]]*\]|\{[^}]*\})*\n?([\s\S]*?)\\end\{\1\}/g, (m, _env: string, _opts: string, body: string) => {
    const language = /language\s*=\s*([\w+#-]+)/i.exec(m)?.[1] ?? /\\begin\{minted\}(?:\[[^\]]*\])?\{([^}]*)\}/.exec(m)?.[1] ?? "";
    kept.push(`\n\n\`\`\`${language.toLowerCase()}\n${body.replace(/\n$/, "")}\n\`\`\`\n\n`);
    return `\u0000${kept.length - 1}\u0000`;
  });
}

/** Commands with one argument turned into Markdown, innermost first so nesting works. */
const WRAPPERS: Record<string, (arg: string) => string> = {
  textbf: (a) => `**${a}**`,
  textit: (a) => `*${a}*`,
  emph: (a) => `*${a}*`,
  textsl: (a) => `*${a}*`,
  underline: (a) => `<u>${a}</u>`,
  uline: (a) => `<u>${a}</u>`,
  texttt: (a) => `\`${a}\``,
  sout: (a) => `~~${a}~~`,
  textsuperscript: (a) => `<sup>${a}</sup>`,
  textsubscript: (a) => `<sub>${a}</sub>`,
  url: (a) => `<${a}>`,
  footnote: (a) => ` (${a})`,
  caption: (a) => `*${a}*`,
  mbox: (a) => a,
  textrm: (a) => a,
  textnormal: (a) => a,
  text: (a) => a,
};

/** Inline commands: wrappers, links, references, escapes. Generated math goes into `kept` (placeholders). */
function inline(text: string, kept: string[]): string {
  const keep = (markdown: string) => {
    kept.push(markdown);
    return `\u0000${kept.length - 1}\u0000`;
  };
  let out = "";
  for (let i = 0; i < text.length; ) {
    if (text[i] !== "\\") {
      out += text[i++];
      continue;
    }
    const name = /^\\([A-Za-z]+)\*?/.exec(text.slice(i));
    if (!name) {
      // \% \& \# \_ \$ \{ \} \\ and the like
      const c = text[i + 1] ?? "";
      if (c === "\\") {
        out += "\\\n";
        i += 2;
      } else if (c === "{" || c === "}") {
        // literal braces survive the removal of grouping braces at the end
        out += c === "{" ? "\u0004" : "\u0005";
        i += 2;
      } else if ("%&".includes(c)) {
        out += c;
        i += 2;
      } else if ("#_$".includes(c)) {
        out += `\\${c}`;
        i += 2;
      } else if (c === ",") {
        out += " ";
        i += 2;
      } else {
        out += c;
        i += 2;
      }
      continue;
    }
    const command = name[1];
    let j = i + name[0].length;
    const wrapper = WRAPPERS[command];
    if (wrapper) {
      const arg = group(text, skipSpaces(text, j));
      if (arg) {
        out += wrapper(inline(arg.content, kept));
        i = arg.end;
        continue;
      }
    }
    if (command === "textcolor" || command === "colorbox") {
      let k = skipSpaces(text, j);
      const model = text[k] === "[" ? /^\[([^\]]*)\]/.exec(text.slice(k)) : null;
      if (model) k = skipSpaces(text, k + model[0].length);
      const color = group(text, k);
      const body = color && group(text, skipSpaces(text, color.end));
      if (color && body) {
        const value = model && /html/i.test(model[1]) ? `#${color.content.trim()}` : color.content.trim();
        out += `<span style="${command === "textcolor" ? "color" : "background-color"}:${value}">${inline(body.content, kept)}</span>`;
        i = body.end;
        continue;
      }
    }
    if (command === "href") {
      const url = group(text, skipSpaces(text, j));
      const label = url && group(text, skipSpaces(text, url.end));
      if (url && label) {
        out += `[${inline(label.content, kept)}](${url.content.replace(/\s/g, "%20")})`;
        i = label.end;
        continue;
      }
    }
    if (command === "ref" || command === "eqref" || command === "autoref" || command === "cref") {
      const arg = group(text, skipSpaces(text, j));
      if (arg) {
        out += keep(`$\\eqref{${arg.content.trim()}}$`);
        i = arg.end;
        continue;
      }
    }
    if (command === "cite" || command === "citep" || command === "citet") {
      j = skipOptional(text, j);
      const arg = group(text, skipSpaces(text, j));
      if (arg) {
        out += `[${arg.content}]`;
        i = arg.end;
        continue;
      }
    }
    if (command === "newline" || command === "linebreak") {
      out += "\\\n";
      i = j;
      continue;
    }
    if (command === "par") {
      out += "\n\n";
      i = j;
      continue;
    }
    if (command === "ldots" || command === "dots") {
      out += "…";
      i = j;
      continue;
    }
    // Commands with arguments that leave nothing (\label, \includegraphics, \vspace, …), and plain switches
    // (\noindent, \centering, \maketitle, …): dropped.
    j = skipOptional(text, j);
    const arg = group(text, skipSpaces(text, j));
    const dropsArgument = /^(label|includegraphics|vspace|hspace|setlength|addtolength|bibliography|bibliographystyle|input|include|usepackage|documentclass|newcommand|renewcommand|pagestyle|thispagestyle|setcounter)$/.test(command);
    i = dropsArgument && arg ? arg.end : j;
    if (!dropsArgument && arg) {
      out += inline(arg.content, kept);
      i = arg.end;
    }
  }
  return out;
}

/** A tabular as a GFM table (the first row is the header). */
function tables(text: string, kept: string[]): string {
  return text.replace(/\\begin\{tabular\}\s*(\{[^}]*\})?([\s\S]*?)\\end\{tabular\}/g, (_m, _spec: string, body: string) => {
    const rows = body
      .replace(/\\(hline|toprule|midrule|bottomrule|cline\{[^}]*\})/g, "")
      .split(/\\\\/)
      .map((row) => row.trim())
      .filter(Boolean)
      .map((row) => row.split(/(?<!\\)&/).map((cell) => inline(cell.trim(), kept).replace(/[{}]/g, "").replace(/\|/g, "\\|").replace(/\n+/g, " ")));
    if (!rows.length) return "";
    const width = Math.max(...rows.map((r) => r.length));
    const line = (cells: string[]) => `| ${Array.from({ length: width }, (_, k) => cells[k] ?? "").join(" | ")} |`;
    kept.push(`\n\n${line(rows[0])}\n| ${Array.from({ length: width }, () => "---").join(" | ")} |\n${rows.slice(1).map(line).join("\n")}\n\n`);
    return `\u0000${kept.length - 1}\u0000`;
  });
}

/** \section{…} and its kin on lines of their own, as headings (\paragraph{…} as a bold run-in). */
function sections(text: string): string {
  const pattern = /\\(chapter|section|subsection|subsubsection|paragraph)\*?\s*(\[[^\]]*\])?\s*(?=\{)/g;
  let out = "";
  let last = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(text))) {
    const title = group(text, match.index + match[0].length);
    if (!title) continue;
    const kind = match[1];
    const marker = kind === "chapter" || kind === "section" ? "#" : kind === "subsection" ? "##" : kind === "subsubsection" ? "###" : null;
    out += text.slice(last, match.index) + (marker ? `\n\n${marker} ${title.content.trim()}\n\n` : `\n\n**${title.content.trim()}** `);
    last = title.end;
    pattern.lastIndex = title.end;
  }
  return out + text.slice(last);
}

/** Lists, quotes and the environments that only arrange text, line by line. */
function blocks(text: string): string {
  const out: string[] = [];
  const lists: Array<"-" | "1."> = [];
  let quote = 0;
  const prefix = () => "> ".repeat(quote);
  for (const raw of text.split("\n")) {
    let line = raw;
    // Several \begin / \end / \item on one line are taken apart first.
    const pieces = line.split(/(?=\\begin\{|\\end\{|\\item\b)/);
    for (let piece of pieces) {
      const begin = /^\\begin\{(itemize|enumerate|description|quote|quotation|center|flushleft|flushright|abstract|figure|table|minipage)\*?\}(\[[^\]]*\])?(\{[^}]*\})?/.exec(piece);
      if (begin) {
        const env = begin[1];
        if (env === "itemize" || env === "description") lists.push("-");
        else if (env === "enumerate") lists.push("1.");
        else if (env === "quote" || env === "quotation") quote++;
        out.push("");
        piece = piece.slice(begin[0].length);
      }
      const end = /^\\end\{(itemize|enumerate|description|quote|quotation|center|flushleft|flushright|abstract|figure|table|minipage)\*?\}/.exec(piece);
      if (end) {
        const env = end[1];
        if (env === "itemize" || env === "description" || env === "enumerate") lists.pop();
        else if (env === "quote" || env === "quotation") quote = Math.max(0, quote - 1);
        out.push("");
        piece = piece.slice(end[0].length);
      }
      const item = /^\\item\b\s*(\[([^\]]*)\])?/.exec(piece);
      if (item && lists.length) {
        const depth = lists.length - 1;
        const indent = lists.slice(0, depth).map((m) => (m === "-" ? "  " : "   ")).join("");
        const label = item[2] ? `**${item[2]}** ` : "";
        out.push(`${prefix()}${indent}${lists[depth]} ${label}${piece.slice(item[0].length).trim()}`);
        continue;
      }
      if (!piece.trim()) continue;
      const indent = lists.length ? lists.map((m) => (m === "-" ? "  " : "   ")).join("") : "";
      out.push(`${prefix()}${indent}${piece.trim()}`);
    }
    if (!raw.trim()) out.push(prefix().trim());
  }
  return out.join("\n");
}

/** LaTeX source as Markdown the document reads. */
export function latexToMarkdown(source: string): string {
  let text = source.replace(/\r\n?/g, "\n");
  // Comments (an unescaped % to the end of the line).
  text = text.replace(/(^|[^\\])%.*$/gm, "$1");
  // The body of a whole document; its title as the first heading.
  const title = /\\title\s*\{([^}]*)\}/.exec(text)?.[1];
  const body = /\\begin\{document\}([\s\S]*?)(\\end\{document\}|$)/.exec(text);
  if (body) text = (title && /\\maketitle/.test(body[1]) ? `\\section{${title}}\n` : "") + body[1];
  const kept: string[] = [];
  text = protectCode(text, kept);
  text = protectMath(text, kept);
  text = tables(text, kept);
  text = sections(text);
  // Old switches in a group: {\bf …}, {\it …}, {\tt …}.
  text = text
    .replace(/\{\\(?:bf|bfseries)\s+([^{}]*)\}/g, "**$1**")
    .replace(/\{\\(?:it|itshape|em|sl|slshape)\s+([^{}]*)\}/g, "*$1*")
    .replace(/\{\\(?:tt|ttfamily)\s+([^{}]*)\}/g, "`$1`");
  text = blocks(text);
  text = inline(text, kept);
  // Grouping braces go; literal ones (\{ \}) come back.
  text = text.replace(/[{}]/g, "").replace(/\u0004/g, "{").replace(/\u0005/g, "}");
  // Typographic leftovers: ~ (a tie), ``quotes'', -- and ---.
  text = text.replace(/(?<!\\)~/g, " ").replace(/``/g, "“").replace(/''/g, "”").replace(/---/g, "—").replace(/--/g, "–");
  for (let pass = 0; pass < 3 && text.includes("\u0000"); pass++) text = text.replace(/\u0000(\d+)\u0000/g, (_m, n: string) => kept[Number(n)]);
  return text.replace(/\n{3,}/g, "\n\n").trim() + "\n";
}
