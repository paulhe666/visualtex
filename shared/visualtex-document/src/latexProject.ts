// A LaTeX project read as one source (a .zip of a paper or a thesis with its class): the main file found, the files
// it \input-s put in place, and the macros its class, packages and preamble define (\newcommand, \def,
// \DeclareMathOperator) applied to the text, so that latexImport.ts reads what TeX would typeset. Theorem-like
// environments (\newtheorem, proof) become a bold numbered label before their text.

export interface ProjectFile {
  /** the path inside the archive ("/" separated) */
  path: string;
  bytes: Uint8Array;
}

export interface LatexProject {
  /** the main file's path */
  main: string;
  /** the main file with its inputs in place and the macros applied */
  source: string;
  /** how many macros were applied (definitions found, used or not) */
  macros: number;
}

/** A file's text: UTF-8, or GB 18030 (CTeX projects) when it is not valid UTF-8. */
export function decodeText(bytes: Uint8Array): string {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    try {
      return new TextDecoder("gb18030").decode(bytes);
    } catch {
      return new TextDecoder("utf-8").decode(bytes);
    }
  }
}

/** Comments: an unescaped % to the end of the line. */
function stripComments(text: string): string {
  return text.replace(/\r\n?/g, "\n").replace(/(^|[^\\])%.*$/gm, "$1");
}

function normalizePath(path: string): string {
  const parts: string[] = [];
  for (const part of path.replace(/\\/g, "/").split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") parts.pop();
    else parts.push(part);
  }
  return parts.join("/");
}

const directoryOf = (path: string) => (path.includes("/") ? path.slice(0, path.lastIndexOf("/") + 1) : "");

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

/** An optional [...] argument at i (brackets may nest braces): its content and the index after it. */
function optional(text: string, i: number): { content: string; end: number } | null {
  if (text[i] !== "[") return null;
  let depth = 0;
  for (let j = i + 1; j < text.length; j++) {
    const c = text[j];
    if (c === "\\") {
      j++;
      continue;
    }
    if (c === "{") depth++;
    else if (c === "}") depth--;
    else if (c === "]" && depth === 0) return { content: text.slice(i + 1, j), end: j + 1 };
  }
  return null;
}

const skipSpace = (text: string, i: number) => {
  while (i < text.length && /\s/.test(text[i])) i++;
  return i;
};

/** One macro argument at i: a {…} group, a control sequence or one character. */
function argument(text: string, i: number): { content: string; end: number } | null {
  i = skipSpace(text, i);
  if (i >= text.length) return null;
  if (text[i] === "{") return group(text, i);
  if (text[i] === "\\") {
    const name = /^\\(?:[A-Za-z]+|.)/.exec(text.slice(i));
    return name ? { content: name[0], end: i + name[0].length } : null;
  }
  return { content: text[i], end: i + 1 };
}

const TEX_EXTENSIONS = ["", ".tex"];

/** The project's files by normalized path; macOS archive leftovers skipped. */
function index(files: ProjectFile[]): Map<string, ProjectFile> {
  const map = new Map<string, ProjectFile>();
  for (const file of files) {
    const path = normalizePath(file.path);
    if (!path || path.startsWith("__MACOSX/") || /(^|\/)\._/.test(path)) continue;
    map.set(path, file);
  }
  return map;
}

/** The main file: the one with \documentclass and a document body, nearest the top, preferring main.tex. */
function findMain(texts: Map<string, string>): string | null {
  const paths = [...texts.keys()].filter((path) => path.toLowerCase().endsWith(".tex"));
  const score = (path: string) => {
    const text = texts.get(path) ?? "";
    let value = 0;
    if (/\\documentclass/.test(text)) value += 100;
    if (/\\begin\{document\}/.test(text)) value += 50;
    if (/\\documentclass\s*\[[^\]]*\bmain\b[^\]]*\]\s*\{subfiles\}/.test(text)) value -= 120;
    if (/^(main|thesis|paper|article|manuscript|report)\.tex$/i.test(path.split("/").pop() ?? "")) value += 10;
    value -= path.split("/").length * 5;
    return value + Math.min(text.length / 100000, 1);
  };
  paths.sort((a, b) => score(b) - score(a));
  return paths[0] ?? null;
}

/** The text of a whole document (a \subfile) without its preamble. */
function bodyOf(text: string): string {
  const body = /\\begin\{document\}([\s\S]*?)(\\end\{document\}|$)/.exec(text);
  return body ? body[1] : text;
}

/** \input, \include, \subfile, \import and \subimport replaced by the files' text (comments removed). */
function inlineInputs(text: string, dir: string, mainDir: string, texts: Map<string, string>, depth: number, seen: Set<string>): string {
  if (depth > 20) return text;
  return text.replace(
    /\\(input|include|subfile|import|subimport|subinputfrom|inputfrom)\b\s*(?:\{([^{}]*)\}\s*)?\{([^{}]*)\}/g,
    (_whole, command: string, first: string | undefined, second: string) => {
      const name = second.trim();
      const candidates: string[] = [];
      if ((command === "import" || command === "inputfrom") && first !== undefined) candidates.push(normalizePath(first + "/" + name));
      else if ((command === "subimport" || command === "subinputfrom") && first !== undefined) candidates.push(normalizePath(dir + first + "/" + name));
      else candidates.push(normalizePath(dir + name), normalizePath(mainDir + name));
      for (const candidate of candidates) {
        for (const extension of TEX_EXTENSIONS) {
          const path = candidate + extension;
          const included = texts.get(path);
          if (included === undefined || seen.has(path)) continue;
          seen.add(path);
          const inner = command === "subfile" ? bodyOf(included) : included;
          const expanded = inlineInputs(inner, directoryOf(path), mainDir, texts, depth + 1, seen);
          seen.delete(path);
          return `\n${expanded}\n`;
        }
      }
      // a file not in the archive: nothing
      return "";
    },
  );
}

interface Macro {
  args: number;
  /** the default of an optional first argument */
  optional?: string;
  body: string;
}

/** Commands latexImport.ts reads itself, or that only TeX internals redefine: never replaced. */
const KEPT = new Set(
  (
    "begin end item section subsection subsubsection chapter paragraph subparagraph part title author date maketitle " +
    "textbf textit emph underline texttt textsc textsf textrm textsl textup textnormal text mbox url href label ref " +
    "eqref autoref cref cite citep citet footnote caption includegraphics newline linebreak par input include " +
    "documentclass usepackage frac dfrac tfrac sqrt left right mathrm mathbf mathit mathsf mathtt mathcal mathbb " +
    "operatorname hline toprule midrule bottomrule tableofcontents bibliography bibliographystyle appendix"
  ).split(" "),
);

/** TeX programming in a definition: such a macro cannot be applied as text. */
const PROGRAMMING = /@|\\(?:if[A-Za-z]*|fi|else|or|expandafter|noexpand|csname|endcsname|def|gdef|edef|xdef|let|futurelet|newcommand|renewcommand|providecommand|global|long|protected|begingroup|endgroup|bgroup|egroup|setlength|addtolength|setcounter|addtocounter|stepcounter|refstepcounter|vskip|hskip|vspace|hspace|vfill|hfill|clearpage|newpage|cleardoublepage|thispagestyle|pagestyle|NewDocumentCommand|ExplSyntaxOn|makeatletter|makeatother|input|include|usepackage|RequirePackage|LoadClass|PassOptionsToClass|ProcessOptions|DeclareOption|newenvironment|renewenvironment)(?![A-Za-z])/;

/** The definitions in a file, in order: \newcommand and its kin, simple \def, \DeclareMathOperator, \newtheorem. */
interface Theorem {
  label: string;
  /** the environment whose counter it uses; null: not numbered */
  counter: string | null;
  /** numbered within chapters or sections ([section]) */
  within?: string;
}

function collectDefinitions(text: string, macros: Map<string, Macro>, theorems: Map<string, Theorem>): void {
  const definition = /\\(newcommand|renewcommand|providecommand|DeclareRobustCommand|DeclareMathOperator|def|newtheorem)(\*?)/g;
  let match: RegExpExecArray | null;
  while ((match = definition.exec(text))) {
    const [, kind, star] = match;
    let i = skipSpace(text, match.index + match[0].length);
    if (kind === "newtheorem") {
      const env = group(text, i);
      if (!env) continue;
      i = skipSpace(text, env.end);
      const shared = optional(text, i);
      if (shared) i = skipSpace(text, shared.end);
      const label = group(text, i);
      if (!label) continue;
      const within = shared ? undefined : optional(text, skipSpace(text, label.end));
      theorems.set(env.content.trim(), {
        label: label.content.trim(),
        counter: star ? null : (shared?.content.trim() ?? env.content.trim()),
        within: within?.content.trim(),
      });
      definition.lastIndex = within?.end ?? label.end;
      continue;
    }
    // the name: {\name} or \name
    let name: string | null = null;
    if (text[i] === "{") {
      const braced = group(text, i);
      const inner = braced && /^\s*\\([A-Za-z]+)\s*$/.exec(braced.content);
      if (!braced || !inner) continue;
      name = inner[1];
      i = braced.end;
    } else {
      const bare = /^\\([A-Za-z]+)/.exec(text.slice(i));
      if (!bare) continue;
      name = bare[1];
      i += bare[0].length;
    }
    let args = 0;
    let defaultValue: string | undefined;
    if (kind === "def") {
      const params = /^((?:#\d)*)\s*(?=\{)/.exec(text.slice(i));
      if (!params) continue; // delimited parameters: TeX programming
      args = params[1].length / 2;
      i += params[0].length;
    } else if (kind !== "DeclareMathOperator") {
      i = skipSpace(text, i);
      const count = optional(text, i);
      if (count) {
        args = Number(count.content.trim()) || 0;
        i = skipSpace(text, count.end);
        const fallback = optional(text, i);
        if (fallback) {
          defaultValue = fallback.content;
          i = skipSpace(text, fallback.end);
        }
      }
    }
    const body = group(text, skipSpace(text, i));
    if (!body) continue;
    definition.lastIndex = body.end;
    if (KEPT.has(name) || args > 9 || body.content.length > 2000) continue;
    if (kind === "DeclareMathOperator") {
      macros.set(name, { args: 0, body: `\\operatorname${star}{${body.content}}` });
      continue;
    }
    if (PROGRAMMING.test(body.content)) continue;
    if (kind === "providecommand" && macros.has(name)) continue;
    macros.set(name, { args, optional: defaultValue, body: body.content });
  }
}

/** Code environments, left as they are while macros are applied. */
const CODE = /\\begin\{(lstlisting|verbatim|minted|comment)\}[\s\S]*?\\end\{\1\}/g;

/** The macros applied to the text, a few passes deep (macros using macros), outside code. */
function applyMacros(text: string, macros: Map<string, Macro>): string {
  if (!macros.size) return text;
  const names = [...macros.keys()].sort((a, b) => b.length - a.length).map((name) => name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  const pattern = new RegExp(`\\\\(${names.join("|")})(?![A-Za-z@])`, "g");
  const limit = Math.max(text.length * 4, 200000);
  const expand = (part: string): string => {
    for (let pass = 0; pass < 8; pass++) {
      let changed = false;
      let out = "";
      let last = 0;
      pattern.lastIndex = 0;
      let match: RegExpExecArray | null;
      while ((match = pattern.exec(part))) {
        // \\R is a line break followed by R
        let slashes = 0;
        for (let k = match.index - 1; k >= 0 && part[k] === "\\"; k--) slashes++;
        if (slashes % 2) continue;
        const macro = macros.get(match[1])!;
        let i = match.index + match[0].length;
        const values: string[] = [];
        let complete = true;
        for (let n = 0; n < macro.args; n++) {
          if (n === 0 && macro.optional !== undefined) {
            const given = optional(part, skipSpace(part, i));
            if (given) {
              values.push(given.content);
              i = given.end;
            } else values.push(macro.optional);
            continue;
          }
          const value = argument(part, i);
          if (!value) {
            complete = false;
            break;
          }
          values.push(value.content);
          i = value.end;
        }
        if (!complete) continue;
        let replacement = macro.body.replace(/#(\d)/g, (_m, k: string) => values[Number(k) - 1] ?? "").replace(/##/g, "#");
        // a control word at the end must not run into the letters after it: \R x, \alpha b
        if (/\\[A-Za-z]+$/.test(replacement) && /^[A-Za-z]/.test(part.slice(i))) replacement += " ";
        out += part.slice(last, match.index) + replacement;
        last = i;
        pattern.lastIndex = i;
        changed = true;
        if (out.length > limit) return part;
      }
      if (!changed) return part;
      part = out + part.slice(last);
    }
    return part;
  };
  let out = "";
  let last = 0;
  for (const code of text.matchAll(CODE)) {
    out += expand(text.slice(last, code.index)) + code[0];
    last = code.index! + code[0].length;
  }
  return out + expand(text.slice(last));
}

/** Theorem-like environments as a bold label ("Theorem 2.1 (Name).") before their text; proof with a closing ∎.
 *  Numbers count per environment (or the one whose counter it shares), within chapters or sections when defined so. */
function applyTheorems(text: string, theorems: Map<string, Theorem>, chinese: boolean): string {
  const counters = new Map<string, number>();
  const sectioning = { chapter: 0, section: 0 };
  const names = [...theorems.keys(), "proof"].map((name) => name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  const pattern = new RegExp(
    `\\\\(chapter|section)(?=\\s*[\\[{])|\\\\begin\\{(${names.join("|")})\\}\\s*(\\[[^\\]]*\\])?\\s*|\\\\end\\{(${names.join("|")})\\}`,
    "g",
  );
  const counterOf = (name: string): string => {
    const shared = theorems.get(name)?.counter;
    return shared && shared !== name && theorems.has(shared) ? counterOf(shared) : name;
  };
  return text.replace(pattern, (whole, unit: string | undefined, begin: string | undefined, note: string | undefined, end: string | undefined) => {
    if (unit === "chapter" || unit === "section") {
      sectioning[unit]++;
      if (unit === "chapter") sectioning.section = 0;
      for (const [name, theorem] of theorems) if (theorem.within === unit || (unit === "chapter" && theorem.within === "section")) counters.delete(name);
      return whole;
    }
    if (end) return end === "proof" ? " $\\square$\n\n" : "\n\n";
    const extra = note ? ` (${note.slice(1, -1).trim()})` : "";
    if (begin === "proof") return `\n\n\\textit{${note ? note.slice(1, -1).trim() : chinese ? "证明" : "Proof"}.} `;
    const theorem = theorems.get(begin!)!;
    let number = "";
    if (theorem.counter) {
      const counter = counterOf(begin!);
      const value = (counters.get(counter) ?? 0) + 1;
      counters.set(counter, value);
      const within = theorems.get(counter)?.within;
      const prefix =
        within === "section" ? `${sectioning.chapter ? `${sectioning.chapter}.` : ""}${sectioning.section}.` : within === "chapter" ? `${sectioning.chapter}.` : "";
      number = ` ${prefix}${value}`;
    }
    return `\n\n\\textbf{${theorem.label}${number}${extra}.} `;
  });
}

// ---------------------------------------------------------------- BibTeX --

interface BibEntry {
  type: string;
  key: string;
  fields: Map<string, string>;
}

/** A .bib file's entries (@string macros applied; @comment / @preamble skipped). */
export function parseBibtex(text: string): BibEntry[] {
  const entries: BibEntry[] = [];
  const strings = new Map<string, string>([["jan", "January"], ["feb", "February"], ["mar", "March"], ["apr", "April"], ["may", "May"], ["jun", "June"], ["jul", "July"], ["aug", "August"], ["sep", "September"], ["oct", "October"], ["nov", "November"], ["dec", "December"]]);
  const at = /@\s*([A-Za-z]+)\s*([{(])/g;
  let match: RegExpExecArray | null;
  while ((match = at.exec(text))) {
    const type = match[1].toLowerCase();
    const close = match[2] === "{" ? "}" : ")";
    // the entry's body: to its closing bracket, braces balanced
    let depth = 0;
    let end = -1;
    for (let i = at.lastIndex; i < text.length; i++) {
      const c = text[i];
      if (c === "{") depth++;
      else if (c === "}" && depth > 0) depth--;
      else if (c === close && depth === 0) {
        end = i;
        break;
      }
    }
    if (end < 0) break;
    const body = text.slice(at.lastIndex, end);
    at.lastIndex = end + 1;
    if (type === "comment" || type === "preamble") continue;
    const fields = bibFields(type === "string" ? `,${body}` : body.slice(body.indexOf(",")), strings);
    if (type === "string") {
      for (const [name, value] of fields) strings.set(name, value);
      continue;
    }
    const key = (body.includes(",") ? body.slice(0, body.indexOf(",")) : body).trim();
    if (key) entries.push({ type, key, fields });
  }
  return entries;
}

/** `, name = {value} # "value" # macro, …` as fields (lower-case names). */
function bibFields(text: string, strings: Map<string, string>): Map<string, string> {
  const fields = new Map<string, string>();
  let i = 0;
  while (i < text.length) {
    const name = /^[\s,]*([A-Za-z][\w:.+-]*)\s*=\s*/.exec(text.slice(i));
    if (!name) break;
    i += name[0].length;
    let value = "";
    for (;;) {
      if (text[i] === "{") {
        const part = group(text, i);
        if (!part) return fields;
        value += part.content;
        i = part.end;
      } else if (text[i] === '"') {
        let k = i + 1;
        let depth = 0;
        while (k < text.length && !(text[k] === '"' && depth === 0)) {
          if (text[k] === "{") depth++;
          else if (text[k] === "}") depth--;
          k++;
        }
        value += text.slice(i + 1, k);
        i = k + 1;
      } else {
        const bare = /^[\w:.+-]+/.exec(text.slice(i));
        if (!bare) break;
        value += strings.get(bare[0].toLowerCase()) ?? bare[0];
        i += bare[0].length;
      }
      const hash = /^\s*#\s*/.exec(text.slice(i));
      if (!hash) break;
      i += hash[0].length;
    }
    fields.set(name[1].toLowerCase(), value.replace(/\s+/g, " ").trim());
  }
  return fields;
}

/** "Last, First and First Last" as "F. Last and F. Last" (a list of more than three: "F. Last et al."). */
function bibAuthors(value: string): string {
  const names = value
    .split(/\s+and\s+/)
    .map((name) => name.trim())
    .filter(Boolean)
    .map((name) => {
      if (/^\{.*\}$/.test(name) || name === "others") return name;
      let first = "";
      let last = name;
      if (name.includes(",")) {
        const parts = name.split(",").map((p) => p.trim());
        last = parts[0];
        first = parts.slice(1).join(" ");
      } else {
        const words = name.split(/\s+/);
        last = words.pop()!;
        first = words.join(" ");
      }
      const initials = first
        .split(/[\s~]+/)
        .filter(Boolean)
        .map((word) => (/[\u4e00-\u9fff]/.test(word) ? word : `${word.replace(/[{}]/g, "").split("-").map((p) => p[0] ?? "").join(".-")}.`))
        .join(" ");
      return initials ? `${initials} ${last}` : last;
    });
  if (names.length > 3) return `${names[0]} et al.`;
  const list = names.map((n) => (n === "others" ? "et al." : n));
  return list.length > 1 ? `${list.slice(0, -1).join(", ")} and ${list[list.length - 1]}` : list[0] ?? "";
}

/** An entry as a reference list item (LaTeX text, as BibTeX's plain style sets it). */
function bibText(entry: BibEntry): string {
  const f = (name: string) => entry.fields.get(name) ?? "";
  const parts: string[] = [];
  const authors = f("author") ? bibAuthors(f("author")) : f("editor") ? `${bibAuthors(f("editor"))}, editors` : "";
  if (authors) parts.push(authors);
  const title = f("title");
  const book = entry.type === "book" || entry.type === "phdthesis" || entry.type === "mastersthesis";
  if (title) parts.push(book ? `\\emph{${title}}` : title);
  const container = f("journal") || f("journaltitle") || f("booktitle");
  if (container) {
    let where = `${entry.type === "inproceedings" || entry.type === "incollection" ? "In " : ""}\\emph{${container}}`;
    if (f("volume")) where += `, ${f("volume")}${f("number") ? `(${f("number")})` : ""}`;
    if (f("pages")) where += `:${f("pages")}`;
    parts.push(where);
  } else if (f("pages")) parts.push(`pp. ${f("pages")}`);
  if (entry.type === "phdthesis") parts.push(`PhD thesis${f("school") ? `, ${f("school")}` : ""}`);
  else if (entry.type === "mastersthesis") parts.push(`Master's thesis${f("school") ? `, ${f("school")}` : ""}`);
  const publisher = f("publisher") || f("institution") || f("organization");
  if (publisher) parts.push(f("address") ? `${publisher}, ${f("address")}` : publisher);
  const year = f("year") || f("date").slice(0, 4);
  if (year) parts.push(f("month") ? `${f("month")} ${year}` : year);
  if (f("note")) parts.push(f("note"));
  const link = f("doi") ? `\\url{https://doi.org/${f("doi").replace(/^https?:\/\/(dx\.)?doi\.org\//, "")}}` : f("url") ? `\\url{${f("url")}}` : "";
  return `${parts.join(". ").replace(/\.\./g, ".")}.${link ? ` ${link}` : ""}`;
}

/**
 * \bibliography{…} / \printbibliography with the project's .bib files: a thebibliography of the cited entries in
 * the order they are first cited (\nocite{*}: all of them), as BibTeX's unsrt style lists them.
 */
function applyBibliography(source: string, bibs: string[]): string {
  if (!bibs.length || !/\\(bibliography|printbibliography)\b/.test(source)) return source;
  const entries = new Map<string, BibEntry>();
  for (const text of bibs) for (const entry of parseBibtex(text)) if (!entries.has(entry.key)) entries.set(entry.key, entry);
  const cited: string[] = [];
  const cite = /\\(?:no)?cite[a-zA-Z]*\*?\s*(?:\[[^\]]*\]\s*){0,2}\{([^{}]*)\}/g;
  for (const match of source.matchAll(cite)) {
    for (const key of match[1].split(",").map((k) => k.trim())) {
      if (key === "*") for (const all of entries.keys()) if (!cited.includes(all)) cited.push(all);
      if (entries.has(key) && !cited.includes(key)) cited.push(key);
    }
  }
  if (!cited.length) return source;
  const list =
    `\\begin{thebibliography}{${String(cited.length).replace(/\d/g, "9")}}\n` +
    cited.map((key) => `\\bibitem{${key}} ${bibText(entries.get(key)!)}`).join("\n") +
    "\n\\end{thebibliography}";
  let placed = false;
  return source.replace(/\\bibliography\s*\{[^{}]*\}|\\printbibliography\b(\s*\[[^\]]*\])?/g, () => {
    if (placed) return "";
    placed = true;
    return list;
  });
}

/** A project archive's files as one LaTeX source for latexToMarkdown; null when it holds no .tex file. */
export function readLatexProject(files: ProjectFile[]): LatexProject | null {
  const byPath = index(files);
  const texts = new Map<string, string>();
  for (const [path, file] of byPath) {
    if (/\.(tex|cls|sty|def|ltx)$/i.test(path)) texts.set(path, stripComments(decodeText(file.bytes)));
  }
  const main = findMain(texts);
  if (!main) return null;
  const mainDir = directoryOf(main);
  const source = inlineInputs(texts.get(main)!, mainDir, mainDir, texts, 0, new Set([main]));
  // definitions: the class and packages of the archive first (only those the document loads, then the rest), then
  // the document's own preamble and text
  const macros = new Map<string, Macro>();
  const theorems = new Map<string, Theorem>();
  const loaded = new Set<string>();
  const loads = /\\(?:documentclass|usepackage|RequirePackage|LoadClass)\s*(?:\[[^\]]*\])?\s*\{([^}]*)\}/g;
  const queue = [source];
  while (queue.length) {
    const text = queue.shift()!;
    for (const load of text.matchAll(loads)) {
      for (const name of load[1].split(",").map((part) => part.trim()).filter(Boolean)) {
        for (const path of [...texts.keys()]) {
          const file = path.split("/").pop()!;
          if ((file === `${name}.cls` || file === `${name}.sty`) && !loaded.has(path)) {
            loaded.add(path);
            queue.push(texts.get(path)!);
          }
        }
      }
    }
  }
  for (const path of loaded) collectDefinitions(texts.get(path)!, macros, theorems);
  collectDefinitions(source, macros, theorems);
  const chinese = /[\u4e00-\u9fff]/.test(source);
  const bibs = [...byPath].filter(([path]) => /\.bib$/i.test(path)).map(([, file]) => decodeText(file.bytes));
  return { main, source: applyBibliography(applyTheorems(applyMacros(source, macros), theorems, chinese), bibs), macros: macros.size };
}

/** A Markdown file of an archive (README.md last), for an archive without LaTeX. */
export function findMarkdownFile(files: ProjectFile[]): ProjectFile | null {
  const markdown = [...index(files).entries()].filter(([path]) => /\.(md|markdown)$/i.test(path));
  markdown.sort(([a], [b]) => Number(/readme/i.test(a)) - Number(/readme/i.test(b)) || a.split("/").length - b.split("/").length);
  return markdown[0]?.[1] ?? null;
}
