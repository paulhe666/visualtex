// Markdown in and out (DESIGN.md §7): the document written as the Markdown that Typora, Obsidian and GitHub read
// (CommonMark + GFM tables, strikethrough and task lists, $…$ / $$…$$ math, ==highlight==, <u> for underline), and
// Markdown — pasted or opened — read back as a document, with what Typora reads besides: footnotes, [toc], x^2^ and
// H~2~O, :emoji:, alerts (> [!NOTE]), headings to level 6, and HTML (shown as it is when the document has no element
// for it). ProseMirror's own Markdown bridge (prosemirror-markdown on markdown-it)
// with VisualTeX's rules for math, highlight and the few HTML tags Markdown writers use.
import MarkdownIt, { type StateBlock, type StateCore, type StateInline, type Token } from "markdown-it";
import { DEFAULT_IMAGE_WIDTH, imageAspect, imageEntry, imageFiles, isImage, isImageLine, resolveImageSource } from "./images.ts";
import { emojiOf } from "./emoji.ts";
import { boardHeight, boardItems, boardRows, canonicalRow } from "./imageBoards.ts";
import { ALERTS } from "./schema.ts";
import { pageCommandLatex, parsePageCommand } from "./pageCommands.ts";
import { MarkdownParser, MarkdownSerializer } from "prosemirror-markdown";
import { Slice, type Mark, type Node, type Schema } from "prosemirror-model";
import { Transform } from "prosemirror-transform";
import { Plugin, PluginKey, type EditorState } from "prosemirror-state";
import { CODE_LANGUAGES } from "./codeBlocks.ts";
import { FONT_SIZES } from "./schema.ts";
import { bibliographyEntries, citeLabel } from "./notes.ts";
import { documentSettings, equationLabels } from "./settings.ts";
import { MARKDOWN_TITLE_CLOSE, MARKDOWN_TITLE_OPEN } from "./latexImport.ts";
import { resolveLatexMarkers } from "./latexMarkers.ts";

/** The highlight colour ==text== gets (the first of the highlight swatches). */
export const MARKDOWN_HIGHLIGHT = "#fff59d";

// ---------------------------------------------------------------- reading --

/** $…$ and \(…\) inline math (Pandoc's rule: no space just inside the dollars, no digit right after the closing one). */
function mathInline(state: StateInline, silent: boolean): boolean {
  const src = state.src;
  const start = state.pos;
  if (src.startsWith("\\(", start)) {
    const end = src.indexOf("\\)", start + 2);
    if (end < 0) return false;
    if (!silent) {
      const token = state.push("math_inline", "", 0);
      token.content = src.slice(start + 2, end).trim();
      token.markup = "\\(";
    }
    state.pos = end + 2;
    return true;
  }
  if (src[start] !== "$" || src[start + 1] === "$") return false;
  if (/\s/.test(src[start + 1] ?? " ")) return false;
  let end = start + 1;
  while ((end = src.indexOf("$", end)) >= 0) {
    if (src[end - 1] === "\\") {
      end += 1;
      continue;
    }
    break;
  }
  if (end < 0 || end === start + 1) return false;
  if (/\s/.test(src[end - 1]) || /\d/.test(src[end + 1] ?? "")) return false;
  if (!silent) {
    const token = state.push("math_inline", "", 0);
    token.content = src.slice(start + 1, end);
    token.markup = "$";
  }
  state.pos = end + 1;
  return true;
}

/** $$ … $$ and \[ … \] blocks (on one line or several). */
function mathBlock(state: StateBlock, startLine: number, endLine: number, silent: boolean): boolean {
  const lineText = (line: number) => state.src.slice(state.bMarks[line] + state.tShift[line], state.eMarks[line]);
  if (state.sCount[startLine] - state.blkIndent >= 4) return false;
  const first = lineText(startLine);
  const open = first.startsWith("$$") ? "$$" : first.startsWith("\\[") ? "\\[" : null;
  if (!open) return false;
  const close = open === "$$" ? "$$" : "\\]";
  let body = first.slice(open.length);
  let line = startLine;
  let found = false;
  const closeAt = body.trimEnd().endsWith(close) && body.trim().length >= close.length;
  if (closeAt) {
    body = body.trimEnd().slice(0, -close.length);
    found = true;
  } else {
    const parts = [body];
    while (++line < endLine) {
      const text = lineText(line);
      if (state.sCount[line] < state.blkIndent && text) break;
      if (text.trimEnd().endsWith(close)) {
        parts.push(text.trimEnd().slice(0, -close.length));
        found = true;
        break;
      }
      parts.push(text);
    }
    body = parts.join("\n");
  }
  if (!found) return false;
  if (silent) return true;
  state.line = line + 1;
  const token = state.push("math_block", "math", 0);
  token.block = true;
  token.content = body.trim();
  token.markup = open;
  token.map = [startLine, state.line];
  return true;
}

/** ==highlight== (Typora, Obsidian). */
function highlight(state: StateInline, silent: boolean): boolean {
  const start = state.pos;
  if (!state.src.startsWith("==", start) || state.src[start + 2] === "=" || /\s/.test(state.src[start + 2] ?? " ")) return false;
  const end = state.src.indexOf("==", start + 2);
  if (end < 0 || end === start + 2 || /\s/.test(state.src[end - 1])) return false;
  if (!silent) {
    state.push("mark_open", "mark", 1).markup = "==";
    const max = state.posMax;
    state.pos = start + 2;
    state.posMax = end;
    state.md.inline.tokenize(state);
    state.posMax = max;
    state.push("mark_close", "mark", -1).markup = "==";
  }
  state.pos = end + 2;
  return true;
}

const HTML_PAIRS: Record<string, string> = {
  u: "u",
  ins: "u",
  mark: "mark",
  s: "s",
  del: "s",
  strike: "s",
  b: "strong",
  strong: "strong",
  i: "em",
  em: "em",
  sup: "superscript",
  sub: "subscript",
  kbd: "kbd",
  code: "code_inline",
};

/** The document's size name nearest to a CSS font size ("1.2em", "18px"; the text is 15px). */
function sizeName(value: string): string | null {
  const match = /^([\d.]+)\s*(em|rem|px|%)?$/.exec(value.trim());
  if (!match) return null;
  const n = Number(match[1]);
  const ratio = match[2] === "px" ? n / 15 : match[2] === "%" ? n / 100 : n;
  if (!(ratio > 0)) return null;
  let best = FONT_SIZES[0];
  for (const size of FONT_SIZES) if (Math.abs(Math.log(size.ratio / ratio)) < Math.abs(Math.log(best.ratio / ratio))) best = size;
  return best.name === "normalsize" ? null : best.name;
}

/** The marks a <span style="..."> stands for: its text colour, background (a highlight) and size. */
function spanMarks(tag: string): Array<{ type: string; attr: string; value: string }> {
  const style = /\bstyle\s*=\s*("([^"]*)"|'([^']*)')/i.exec(tag);
  const marks: Array<{ type: string; attr: string; value: string }> = [];
  for (const declaration of (style?.[2] ?? style?.[3] ?? "").split(";")) {
    const [property, ...rest] = declaration.split(":");
    const value = rest.join(":").trim();
    if (!property || !value) continue;
    const name = property.trim().toLowerCase();
    if (name === "color") marks.push({ type: "textColor", attr: "color", value });
    else if (name === "background-color" || name === "background") marks.push({ type: "bgcolor", attr: "color", value });
    else if (name === "font-size") {
      const size = sizeName(value);
      if (size) marks.push({ type: "fontSize", attr: "size", value: size });
    }
  }
  return marks;
}

/**
 * After parsing: the HTML tags Markdown writers use for what Markdown lacks become marks (<u>, <mark>, <s>, <br>);
 * other HTML, images and footnote syntax stay as text (nothing is lost); table cells get a paragraph.
 */
function tidy(state: StateCore) {
  const tokens = state.tokens;
  const Token = state.Token;
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token.type === "html_block") {
      // <br clear="all"> (written after a wrapped image's text): the layout's, not text
      if (/^\s*<br\s+clear\s*=\s*["']?\w+["']?\s*\/?>\s*$/i.test(token.content)) {
        tokens.splice(i, 1);
        i--;
        continue;
      }
      // a page command: <!-- \pagenumbering{roman} -->, or a page break as HTML writes it
      const page = pageCommandToken(state, token.content);
      if (page) {
        tokens.splice(i, 1, page);
        continue;
      }
      const board = boardTokens(state, token.content);
      if (board) {
        tokens.splice(i, 1, ...board);
        i += board.length - 1;
        continue;
      }
      const line = imageLineTokens(state, token.content);
      if (line) {
        tokens.splice(i, 1, ...line);
        i += line.length - 1;
        continue;
      }
      // any other HTML (<details>, <div>, a comment…): kept and shown as it is; an element left open (<details>
      // with Markdown inside, after a blank line) takes in what follows up to the HTML that closes it, as the
      // browser does when GitHub and Typora show it
      let last = i;
      const open = openElements(token.content);
      if (open > 0 && token.level === 0 && token.map) {
        let depth = open;
        for (let k = i + 1; k < tokens.length && depth > 0; k++) {
          if (tokens[k].type === "html_block" && tokens[k].level === 0) {
            depth += openElements(tokens[k].content);
            if (depth <= 0) last = k;
          }
        }
      }
      const html = new Token("vtx_html", "div", 0);
      html.block = true;
      html.map = token.map;
      const lines = state.src.split("\n");
      html.attrSet("html", (last > i && token.map && tokens[last].map ? lines.slice(token.map[0], tokens[last].map![1]).join("\n") : token.content).replace(/\n+$/, ""));
      tokens.splice(i, last - i + 1, html);
      continue;
    }
    if ((token.type === "td_open" || token.type === "th_open") && tokens[i + 1]?.type === "inline") {
      tokens.splice(i + 1, 0, new Token("paragraph_open", "p", 1));
      tokens.splice(i + 3, 0, new Token("paragraph_close", "p", -1));
      i += 3;
      continue;
    }
    if (token.type !== "inline" || !token.children) continue;
    const children: Token[] = [];
    // The marks each open <span> stood for (none: the span stays text, and so does its </span>); so for <font>, <a>.
    const spans: Array<string[] | null> = [];
    const fonts: boolean[] = [];
    const anchors: boolean[] = [];
    let skipUntil = -1;
    for (const [index, child] of token.children.entries()) {
      if (index <= skipUntil) continue;
      if (child.type === "html_inline") {
        const tag = /^<\s*(\/?)\s*([a-zA-Z]+)[^>]*?(\/?)\s*>$/.exec(child.content);
        const name = tag?.[2].toLowerCase();
        if (name === "br") {
          if (/\sclear\s*=/i.test(child.content)) continue;
          children.push(new Token("hardbreak", "br", 0));
          continue;
        }
        if (name === "span" && !tag![3]) {
          if (!tag![1]) {
            const marks = spanMarks(child.content);
            spans.push(marks.length ? marks.map((m) => m.type) : null);
            if (marks.length) {
              for (const mark of marks) {
                const open = new Token(`${mark.type}_open`, "span", 1);
                open.attrSet(mark.attr, mark.value);
                children.push(open);
              }
              continue;
            }
          } else {
            const marks = spans.pop();
            if (marks) {
              for (const type of [...marks].reverse()) children.push(new Token(`${type}_close`, "span", -1));
              continue;
            }
          }
        }
        if (name && HTML_PAIRS[name] && !tag![3]) {
          const type = HTML_PAIRS[name];
          // <code>…</code>: its text as inline code
          if (type === "code_inline") {
            if (tag![1]) continue;
            const end = token.children!.findIndex((c, k) => k > index && c.type === "html_inline" && /^<\/\s*code\s*>$/i.test(c.content));
            if (end > 0) {
              const code = new Token("code_inline", "code", 0);
              code.content = token.children!.slice(index + 1, end).map((c) => c.content).join("");
              children.push(code);
              skipUntil = end;
              continue;
            }
          } else {
            children.push(new Token(`${type}_${tag![1] ? "close" : "open"}`, name, tag![1] ? -1 : 1));
            continue;
          }
        }
        // <font color="…"> (old HTML): the text's colour; <a href="…">: a link
        if (name === "font" && !tag![3]) {
          if (!tag![1]) {
            const color = tagAttr(child.content, "color");
            fonts.push(!!color);
            if (color) {
              const open = new Token("textColor_open", "span", 1);
              open.attrSet("color", color);
              children.push(open);
              continue;
            }
          } else if (fonts.pop()) {
            children.push(new Token("textColor_close", "span", -1));
            continue;
          }
        }
        if (name === "a" && !tag![3]) {
          if (!tag![1]) {
            const href = tagAttr(child.content, "href");
            anchors.push(!!href);
            if (href) {
              const open = new Token("link_open", "a", 1);
              open.attrSet("href", href);
              const title = tagAttr(child.content, "title");
              if (title) open.attrSet("title", title);
              children.push(open);
              continue;
            }
          } else if (anchors.pop()) {
            children.push(new Token("link_close", "a", -1));
            continue;
          }
        }
      }
      if (child.type === "image" || (child.type === "html_inline" && /^<img\b/i.test(child.content))) {
        const image = child.type === "image" ? markdownImageToken(state, child) : htmlImageToken(state, child.content);
        if (image) {
          children.push(image);
          // the space written after a wrapped image (its paragraph's line) is not text
          const next = token.children[index + 1];
          if (image.attrGet("wrap") && next?.type === "text") next.content = next.content.replace(/^\s+/, "");
          continue;
        }
      }
      // a picture that could not be found: kept as written (drawn from its address)
      if (child.type === "image" || (child.type === "html_inline" && /^<img\b/i.test(child.content))) {
        const kept = new Token("vtx_html_inline", "span", 0);
        kept.attrSet("html", child.type === "image" ? imageSource(child) : child.content);
        children.push(kept);
        continue;
      }
      if (child.type === "html_inline" || child.type === "image") {
        const text = new Token("text", "", 0);
        text.content = child.type === "image" ? `![${child.content}](${child.attrGet("src") ?? ""})` : child.content;
        children.push(text);
        continue;
      }
      children.push(child);
    }
    token.children = children;
  }
}

/** A Markdown picture as written: ![alt](src "title"). */
function imageSource(token: Token): string {
  const alt = token.children?.map((c) => c.content).join("") ?? token.content;
  const title = token.attrGet("title");
  const src = token.attrGet("src") ?? "";
  return `![${alt}](${/[\s()]/.test(src) ? `<${src}>` : src}${title ? ` "${title.replace(/"/g, '\\"')}"` : ""})`;
}

/** How many more elements an HTML text opens than it closes (void elements and comments aside). */
function openElements(html: string): number {
  const text = html.replace(/<!--[\s\S]*?-->/g, "");
  let depth = 0;
  for (const tag of text.matchAll(/<(\/?)([a-zA-Z][\w-]*)\b[^>]*?(\/?)>/g)) {
    if (tag[3] || /^(br|hr|img|input|meta|link|col|area|base|wbr|source|track|embed|param)$/i.test(tag[2])) continue;
    depth += tag[1] ? -1 : 1;
  }
  return depth;
}

/** An attribute of an HTML tag's text. */
function tagAttr(tag: string, name: string): string | null {
  const match = new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+))`, "i").exec(tag);
  return match ? (match[1] ?? match[2] ?? match[3] ?? "") : null;
}

/** A width as a fraction of the line: "40%", or pixels at 96 per inch on an A4 page's text width. */
function widthFraction(value: string | null): number | null {
  if (!value) return null;
  const match = /^\s*([\d.]+)\s*(%|px)?\s*$/.exec(value);
  if (!match) return null;
  const n = Number(match[1]);
  if (!(n > 0)) return null;
  const fraction = match[2] === "%" ? n / 100 : n / 605;
  return Math.max(0.01, Math.min(1, Math.round(fraction * 1000) / 1000));
}

/** An image token for the parser (the image node), or null when its file is not known. */
function imageToken(state: StateCore, src: string, alt: string, width: number | null, align: string | null): Token | null {
  const key = resolveImageSource(src);
  if (!key) return null;
  const token = new state.Token("vtx_image", "img", 0);
  token.attrSet("src", key);
  token.attrSet("alt", alt);
  // no width given: the picture's own size, as Typora and the browsers show it (on an A4 page's text width)
  const natural = imageEntry(key)?.width ?? 0;
  token.attrSet("width", String(width ?? (natural > 0 ? Math.max(0.05, Math.min(1, Math.round((natural / 605) * 1000) / 1000)) : DEFAULT_IMAGE_WIDTH)));
  if (align === "left" || align === "right") token.attrSet("wrap", align);
  return token;
}

function markdownImageToken(state: StateCore, child: Token): Token | null {
  const alt = child.children?.map((c) => c.content).join("") ?? child.content;
  return imageToken(state, child.attrGet("src") ?? "", alt, null, null);
}

function htmlImageToken(state: StateCore, tag: string, align: string | null = tagAttr(tag, "align")): Token | null {
  const token = imageToken(state, tagAttr(tag, "src") ?? "", tagAttr(tag, "alt") ?? "", widthFraction(tagAttr(tag, "width") ?? cssWidth(tag)), align);
  // Typora's <img style="zoom:50%">: half the picture's own width (on an A4 page's text width)
  const zoom = /zoom\s*:\s*([\d.]+)\s*(%?)/i.exec(tagAttr(tag, "style") ?? "");
  if (token && zoom && !tagAttr(tag, "width")) {
    const factor = Number(zoom[1]) / (zoom[2] ? 100 : 1);
    const natural = imageEntry(token.attrGet("src") ?? "")?.width ?? 0;
    if (factor > 0 && natural > 0) token.attrSet("width", String(Math.max(0.01, Math.min(1, Math.round(((natural * factor) / 605) * 1000) / 1000))));
  }
  return token;
}

/** <img style="width:…">'s width. */
function cssWidth(tag: string): string | null {
  return /(?:^|;)\s*width\s*:\s*([\d.]+\s*(?:%|px)?)/i.exec(tagAttr(tag, "style") ?? "")?.[1] ?? null;
}

/** `<p align="…"><img …> <img …></p>` (VisualTeX's image line) as a paragraph of images. */
function imageLineTokens(state: StateCore, html: string): Token[] | null {
  // also <img> tags alone on their lines (an HTML block): a left-aligned image line
  const match =
    /^\s*<p(\s[^>]*)?>\s*((?:<img\b[^>]*>(?:\s|&[a-z]+;)*)+)<\/p>\s*$/i.exec(html) ?? /^()\s*((?:<img\b[^>]*>(?:\s|&[a-z]+;)*)+)$/i.exec(html);
  if (!match) return null;
  const images: Token[] = [];
  for (const tag of match[2].match(/<img\b[^>]*>/gi) ?? []) {
    const image = htmlImageToken(state, tag, null);
    if (!image) return null;
    images.push(image);
  }
  const open = new state.Token("paragraph_open", "p", 1);
  const align = tagAttr(match[1] ?? "", "align");
  if (align === "center" || align === "right") open.attrSet("align", align);
  const inline = new state.Token("inline", "", 0);
  inline.children = images;
  inline.content = "";
  return [open, inline, new state.Token("paragraph_close", "p", -1)];
}

/**
 * An image board as VisualTeX writes it (<div data-vtx-board> with images carrying data-x and data-top, or, from
 * LaTeX's picture, data-y above the bottom of a board data-height high): the board, or null (then it is kept as HTML).
 */
function boardTokens(state: StateCore, html: string): Token[] | null {
  const open = /^\s*<div\b[^>]*\bdata-vtx-board\b[^>]*>/i.exec(html);
  if (!open) return null;
  const height = Number(tagAttr(open[0], "data-height"));
  const images: Token[] = [];
  for (const tag of html.match(/<img\b[^>]*>/gi) ?? []) {
    const key = resolveImageSource(tagAttr(tag, "src") ?? "");
    const width = widthFraction(tagAttr(tag, "width"));
    const x = Number(tagAttr(tag, "data-x"));
    if (!key || !width || !Number.isFinite(x)) return null;
    let top = Number(tagAttr(tag, "data-top"));
    const y = tagAttr(tag, "data-y");
    if (y !== null && Number.isFinite(height)) top = height - Number(y) - width / imageAspect(key);
    if (!Number.isFinite(top)) return null;
    const token = new state.Token("vtx_board_image", "img", 0);
    token.attrSet("src", key);
    token.attrSet("alt", tagAttr(tag, "alt") ?? "");
    token.attrSet("width", String(width));
    token.attrSet("x", String(x));
    token.attrSet("top", String(Math.round(top * 1000) / 1000));
    images.push(token);
  }
  if (!images.length) return null;
  const start = new state.Token("vtx_board_open", "div", 1);
  start.block = true;
  const end = new state.Token("vtx_board_close", "div", -1);
  end.block = true;
  return [start, ...images, end];
}

/** A page command block: <!-- \newpage --> and the like, or <div style="page-break-after: always"></div>. */
function pageCommandToken(state: StateCore, html: string): Token | null {
  const comment = /^\s*<!--\s*(\\[\s\S]*?)\s*-->\s*$/.exec(html);
  const parsed = comment ? parsePageCommand(comment[1]) : /^\s*<div\s+style\s*=\s*["']\s*page-break-(?:after|before)\s*:\s*always;?\s*["']\s*>\s*<\/div>\s*$/i.test(html) ? { command: "newpage", value: "" } : null;
  if (!parsed) return null;
  const token = new state.Token("vtx_page_command", "div", 0);
  token.block = true;
  token.attrSet("command", parsed.command);
  token.attrSet("value", parsed.value);
  return token;
}

/** <h1 align="center">…</h1> on lines of its own: the document's title, its text read as Markdown. */
function titleLines(state: StateCore) {
  const tokens = state.tokens;
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token.type !== "html_block" || token.level !== 0) continue;
    const match = /^<h1\s+align\s*=\s*["']?center["']?\s*>([\s\S]*?)<\/h1>\s*$/i.exec(token.content);
    if (!match) continue;
    const open = new state.Token("title_open", "p", 1);
    open.map = token.map;
    open.block = true;
    const inline = new state.Token("inline", "", 0);
    inline.content = match[1].replace(/\s*\n\s*/g, " ").trim();
    inline.map = token.map;
    inline.children = [];
    const close = new state.Token("title_close", "p", -1);
    close.block = true;
    tokens.splice(i, 1, open, inline, close);
    i += 2;
  }
}

/** <p align="center|right">…</p> on lines of its own (Typora's, and how aligned paragraphs are written): an aligned
 *  paragraph, its text read as Markdown (an image line is the image rule's). */
function alignedParagraphs(state: StateCore) {
  const tokens = state.tokens;
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token.type !== "html_block") continue;
    // also <div align="…"> and <center> (old HTML Typora still shows)
    const match =
      /^\s*<(p|div)\s+align\s*=\s*["']?(center|right|left)["']?\s*>([\s\S]*?)<\/\1>\s*$/i.exec(token.content)?.slice(1) ??
      /^\s*<(center)>([\s\S]*?)<\/center>\s*$/i.exec(token.content)?.slice(0, 3).map((part, k) => (k === 0 ? "center" : k === 1 ? "center" : part));
    if (!match || /<img\b/i.test(match[2]) || /<(p|div|table|ul|ol|h\d)\b/i.test(match[2])) continue;
    const open = new state.Token("paragraph_open", "p", 1);
    open.map = token.map;
    open.block = true;
    const align = match[1].toLowerCase();
    if (align !== "left") open.attrSet("align", align);
    const inline = new state.Token("inline", "", 0);
    inline.content = match[2].replace(/\s*\n\s*/g, " ").trim();
    inline.map = token.map;
    inline.children = [];
    const close = new state.Token("paragraph_close", "p", -1);
    close.block = true;
    tokens.splice(i, 1, open, inline, close);
    i += 2;
  }
}

/** x^2^ and H~2~O (Typora's superscript and subscript, markdown-it-sup/-sub's rule): no unescaped space inside; "~~"
 *  stays strikethrough. */
function script(marker: "^" | "~", name: "superscript" | "subscript") {
  return (state: StateInline, silent: boolean): boolean => {
    const start = state.pos;
    const max = state.posMax;
    if (silent || state.src[start] !== marker || state.src[start + 1] === marker || start + 2 >= max) return false;
    let end = start + 1;
    while (end < max && state.src[end] !== marker) end += state.src[end] === "\\" ? 2 : 1;
    if (end >= max || end === start + 1) return false;
    const content = state.src.slice(start + 1, end);
    if (/(^|[^\\])(\\\\)*\s/.test(content)) return false;
    state.posMax = end;
    state.pos = start + 1;
    state.push(`${name}_open`, marker === "^" ? "sup" : "sub", 1).markup = marker;
    state.md.inline.tokenize(state);
    state.push(`${name}_close`, marker === "^" ? "sup" : "sub", -1).markup = marker;
    state.pos = end + 1;
    state.posMax = max;
    return true;
  };
}

/** :smile: as 😄 (GitHub's short codes). */
function emoji(state: StateInline, silent: boolean): boolean {
  if (state.src[state.pos] !== ":") return false;
  const match = /^:([a-z0-9_+-]{1,40}):/.exec(state.src.slice(state.pos, state.pos + 44));
  const found = match ? emojiOf(match[1]) : undefined;
  if (!match || !found) return false;
  if (!silent) state.push("text", "", 0).content = found;
  state.pos += match[0].length;
  return true;
}

/** A footnote's Markdown as the footnote's text (its formulas kept; emphasis, code and link marks dropped). */
function footnoteText(markdown: string): string {
  return markdown
    .replace(/\s*\n\s*/g, " ")
    .replace(/!?\[([^\]]*)\]\([^)]*\)/g, "$1")
    .replace(/(\*\*|__|~~|==)(.+?)\1/g, "$2")
    .replace(/(^|[^\\*])\*(?!\s)([^*]+?)\*/g, "$1$2")
    .replace(/`([^`]+)`/g, "$1")
    .trim();
}

/** [^label]: text (with lines indented under it): a footnote's definition, kept for its references. */
function footnoteDefinition(state: StateBlock, startLine: number, endLine: number, silent: boolean): boolean {
  if (state.sCount[startLine] - state.blkIndent >= 4) return false;
  const start = state.bMarks[startLine] + state.tShift[startLine];
  const first = /^\[\^([^\]\s]+)\]:[ \t]?/.exec(state.src.slice(start, state.eMarks[startLine]));
  if (!first) return false;
  if (silent) return true;
  const lines = [state.src.slice(start + first[0].length, state.eMarks[startLine])];
  let line = startLine + 1;
  for (; line < endLine; line++) {
    const text = state.src.slice(state.bMarks[line] + state.tShift[line], state.eMarks[line]);
    // blank: the definition goes on only with an indented line after it
    if (!text.trim()) {
      if (line + 1 < endLine && state.sCount[line + 1] >= state.blkIndent + 2 && state.src.slice(state.bMarks[line + 1], state.eMarks[line + 1]).trim()) continue;
      break;
    }
    if (state.sCount[line] < state.blkIndent + 2 && (/^\[\^[^\]\s]+\]:/.test(text) || /^(#{1,6}\s|[-*+]\s|\d+[.)]\s|>|```|~~~|\$\$|\|)/.test(text))) break;
    lines.push(text);
  }
  const env = state.env as { vtxFootnotes?: Record<string, string> };
  env.vtxFootnotes ??= {};
  env.vtxFootnotes[first[1]] ??= footnoteText(lines.join("\n"));
  state.line = line;
  return true;
}

/** [^label]: a footnote with its definition's text. */
function footnoteReference(state: StateInline, silent: boolean): boolean {
  if (!state.src.startsWith("[^", state.pos)) return false;
  const match = /^\[\^([^\]\s]+)\]/.exec(state.src.slice(state.pos, state.pos + 200));
  const text = match ? (state.env as { vtxFootnotes?: Record<string, string> }).vtxFootnotes?.[match[1]] : undefined;
  if (!match || text === undefined) return false;
  if (!silent) state.push("vtx_footnote", "", 0).attrSet("text", text);
  state.pos += match[0].length;
  return true;
}

/**
 * Before the inline text is read: "[ ] " / "[x] " opening a list item (a task), "> [!NOTE]" opening a quote (an
 * alert, its kind on the quote) and a paragraph of just [toc] (the table of contents).
 */
function blockMarkers(state: StateCore) {
  const tokens = state.tokens;
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    const inline = tokens[i + 2];
    if (token.type === "list_item_open" && tokens[i + 1]?.type === "paragraph_open" && inline?.type === "inline") {
      const task = /^\[([ xX])\](?:[ \t]+|$)/.exec(inline.content);
      if (task) {
        token.attrSet("checked", task[1] === " " ? "false" : "true");
        inline.content = inline.content.slice(task[0].length);
      }
    }
    if (token.type === "blockquote_open" && tokens[i + 1]?.type === "paragraph_open" && inline?.type === "inline") {
      const alert = /^\[!(\w+)\][ \t]*(?:\n|$)/.exec(inline.content);
      if (alert && ALERTS.includes(alert[1].toLowerCase())) {
        token.attrSet("alert", alert[1].toLowerCase());
        inline.content = inline.content.slice(alert[0].length);
        // the marker alone on its line: no paragraph of its own
        if (!inline.content.trim()) {
          tokens.splice(i + 1, 3);
        }
      }
    }
    if (token.type === "paragraph_open" && inline?.type === "paragraph_close" && /^\s*\[toc\]\s*$/i.test(tokens[i + 1]?.content ?? "")) {
      const toc = new state.Token("vtx_toc", "div", 0);
      toc.block = true;
      toc.map = token.map;
      tokens.splice(i, 3, toc);
    }
  }
}

function createMarkdownIt(): MarkdownIt {
  const md = new MarkdownIt("default", { html: true, linkify: true, breaks: false });
  md.block.ruler.before("fence", "math_block", mathBlock, { alt: ["paragraph", "reference", "blockquote", "list"] });
  md.inline.ruler.before("escape", "math_inline", mathInline);
  md.inline.ruler.before("emphasis", "highlight", highlight);
  md.inline.ruler.after("emphasis", "superscript", script("^", "superscript"));
  md.inline.ruler.after("emphasis", "subscript", script("~", "subscript"));
  md.inline.ruler.after("emphasis", "emoji", emoji);
  md.inline.ruler.before("link", "footnote_reference", footnoteReference);
  md.block.ruler.before("reference", "footnote_definition", footnoteDefinition);
  md.core.ruler.before("inline", "vtx_block_markers", blockMarkers);
  md.core.ruler.before("inline", "vtx_title", titleLines);
  md.core.ruler.before("inline", "vtx_aligned", alignedParagraphs);
  md.core.ruler.push("vtx_tidy", tidy);
  return md;
}

/** A fence's language as the block's language ("py" -> python), or "" when the block does not offer it. */
export function blockLanguage(info: string): string {
  const name = info.trim().split(/\s+/)[0]?.toLowerCase() ?? "";
  const alias: Record<string, string> = { py: "python", js: "javascript", ts: "typescript", "c++": "cpp", cs: "csharp", sh: "bash", shell: "bash", zsh: "bash", ps1: "powershell", yml: "yaml", tex: "latex", md: "markdown", golang: "go", rs: "rust", kt: "kotlin", rb: "ruby", m: "matlab", jl: "julia", f90: "fortran", hs: "haskell" };
  const id = alias[name] ?? name;
  return CODE_LANGUAGES.some((l) => l.id === id) ? id : "";
}

/**
 * A display formula's LaTeX: \tag{…} / \label{…} are the document's business (it numbers its formulas). A label
 * becomes the formula's id, so references to it (`$\eqref{label}$`) find it and it is written back as it was.
 */
function displayAttrs(latex: string): { latex: string; numbered: boolean; id: string | null } {
  const label = /\\label\{([^}]*)\}/.exec(latex)?.[1]?.trim() || null;
  const numbered = /\\tag\*?\{[^}]*\}/.test(latex) || !!label;
  const clean = latex.replace(/\\tag\*?\{[^}]*\}/g, "").replace(/\\label\{[^}]*\}/g, "").replace(/\\nonumber|\\notag/g, "").trim();
  const environment = /^\\begin\{(equation|align|gather)\}([\s\S]*)\\end\{\1\}$/.exec(clean);
  if (environment) {
    const inner = environment[2].trim();
    return { latex: environment[1] === "equation" ? inner : `\\begin{${environment[1] === "align" ? "aligned" : "gathered"}}${inner}\\end{${environment[1] === "align" ? "aligned" : "gathered"}}`, numbered: true, id: label };
  }
  return { latex: clean, numbered, id: label };
}

const REFERENCE = /^\s*\\(?:eq)?ref\{([^}]+)\}\s*$/;

/** Formulas that are only a reference to a labelled display formula (`$\eqref{label}$`) become references. */
function linkReferences(doc: Node): Node {
  const labels = new Set<string>();
  doc.descendants((node) => {
    if (node.type.name === "math_display" && node.attrs.id) labels.add(node.attrs.id);
    return true;
  });
  if (!labels.size) return doc;
  const tr = new Transform(doc);
  const found: Array<{ pos: number; size: number; target: string }> = [];
  doc.descendants((node, pos) => {
    const label = node.type.name === "math_inline" ? REFERENCE.exec(node.attrs.latex)?.[1]?.trim() : undefined;
    if (label && labels.has(label)) found.push({ pos, size: node.nodeSize, target: label });
    return true;
  });
  for (const ref of found.reverse()) tr.replaceWith(ref.pos, ref.pos + ref.size, doc.type.schema.nodes.eqref.create({ target: ref.target }));
  return tr.doc;
}

/** A GFM table column's alignment (markdown-it writes it as the cell's style). */
function cellAlign(tok: Token): { align: string | null } {
  return { align: /text-align\s*:\s*(left|center|right)/.exec(tok.attrGet("style") ?? "")?.[1] ?? null };
}

const parsers = new WeakMap<Schema, MarkdownParser>();

function parserFor(schema: Schema): MarkdownParser {
  let parser = parsers.get(schema);
  if (parser) return parser;
  parser = new MarkdownParser(schema, createMarkdownIt(), {
    blockquote: { block: "blockquote", getAttrs: (tok) => ({ alert: tok.attrGet("alert") }) },
    vtx_toc: { node: "toc", noCloseToken: true },
    vtx_board: { block: "image_board" },
    vtx_board_image: {
      node: "board_image",
      getAttrs: (tok) => ({ src: tok.attrGet("src"), alt: tok.attrGet("alt") ?? "", width: Number(tok.attrGet("width")), x: Number(tok.attrGet("x")), top: Number(tok.attrGet("top")) }),
      noCloseToken: true,
    },
    vtx_html_inline: { node: "html_inline", getAttrs: (tok) => ({ html: tok.attrGet("html") ?? "" }), noCloseToken: true },
    vtx_html: { node: "html_block", getAttrs: (tok) => ({ html: tok.attrGet("html") ?? "" }), noCloseToken: true },
    vtx_footnote: { node: "footnote", getAttrs: (tok) => ({ text: tok.attrGet("text") ?? "" }), noCloseToken: true },
    paragraph: { block: "paragraph", getAttrs: (tok) => (tok.attrGet("align") ? { align: tok.attrGet("align") } : null) },
    vtx_page_command: { node: "page_command", getAttrs: (tok) => ({ command: tok.attrGet("command"), value: tok.attrGet("value") ?? "" }), noCloseToken: true },
    vtx_image: {
      node: "image",
      getAttrs: (tok) => ({ src: tok.attrGet("src"), alt: tok.attrGet("alt") ?? "", width: Number(tok.attrGet("width")) || DEFAULT_IMAGE_WIDTH, wrap: tok.attrGet("wrap") }),
      noCloseToken: true,
    },
    list_item: { block: "list_item", getAttrs: (tok) => ({ checked: tok.attrGet("checked") === "true" ? true : tok.attrGet("checked") === "false" ? false : null }) },
    bullet_list: { block: "bullet_list" },
    ordered_list: { block: "ordered_list", getAttrs: (tok) => ({ start: Number(tok.attrGet("start")) || 1 }) },
    heading: { block: "heading", getAttrs: (tok) => ({ level: Math.min(6, Number(tok.tag.slice(1)) || 1) }) },
    title: { block: "title" },
    code_block: { block: "code_block", noCloseToken: true },
    fence: { block: "code_block", getAttrs: (tok) => ({ language: blockLanguage(tok.info) }), noCloseToken: true },
    hr: { node: "horizontal_rule" },
    hardbreak: { node: "hard_break" },
    table: { block: "table" },
    thead: { ignore: true },
    tbody: { ignore: true },
    tr: { block: "table_row" },
    th: { block: "table_cell", getAttrs: cellAlign },
    td: { block: "table_cell", getAttrs: cellAlign },
    math_inline: { node: "math_inline", getAttrs: (tok) => ({ latex: tok.content }), noCloseToken: true },
    math_block: { node: "math_display", getAttrs: (tok) => displayAttrs(tok.content), noCloseToken: true },
    em: { mark: "em" },
    strong: { mark: "strong" },
    s: { mark: "strike" },
    u: { mark: "underline" },
    superscript: { mark: "superscript" },
    subscript: { mark: "subscript" },
    kbd: { mark: "kbd" },
    mark: { mark: "highlight", attrs: { color: MARKDOWN_HIGHLIGHT } },
    textColor: { mark: "textColor", getAttrs: (tok) => ({ color: tok.attrGet("color") }) },
    bgcolor: { mark: "highlight", getAttrs: (tok) => ({ color: tok.attrGet("color") }) },
    fontSize: { mark: "fontSize", getAttrs: (tok) => ({ size: tok.attrGet("size") }) },
    link: { mark: "link", getAttrs: (tok) => ({ href: tok.attrGet("href"), title: tok.attrGet("title") || null }) },
    code_inline: { mark: "code", noCloseToken: true },
  });
  parsers.set(schema, parser);
  return parser;
}

/** Markdown text as a document. */
export function markdownToDoc(schema: Schema, text: string, citationKeys?: Map<string, string>): Node {
  return linkReferences(resolveLatexMarkers(parserFor(schema).parse(text.replace(/\r\n?/g, "\n")), citationKeys));
}

/** Markdown text as a slice to paste (a single paragraph pastes inline); `citationKeys`: the document's entries. */
export function markdownToSlice(schema: Schema, text: string, citationKeys?: Map<string, string>): Slice {
  const doc = markdownToDoc(schema, text, citationKeys);
  const single = doc.childCount === 1 && doc.firstChild!.type.name === "paragraph";
  if (single) return new Slice(doc.firstChild!.content, 0, 0);
  // open into the first and last blocks only when they hold text (a page command or a rule is closed)
  return new Slice(doc.content, doc.firstChild?.isTextblock ? 1 : 0, doc.lastChild?.isTextblock ? 1 : 0);
}

/** Whether pasted text reads as Markdown (blocks or marks), not just plain prose. */
export function looksLikeMarkdown(text: string): boolean {
  return (
    /^(#{1,6}\s|\s*[-*+]\s|\s*\d+[.)]\s|>\s?|```|~~~|\|.*\|\s*$|\$\$|---\s*$)/m.test(text) ||
    /\*\*[^*\n]+\*\*|__[^_\n]+__|~~[^~\n]+~~|==[^=\n]+==|`[^`\n]+`|\[[^\]\n]+\]\([^)\n]+\)/.test(text)
  );
}

/**
 * Whether pasted text is unmistakably a Markdown document (an AI answer, a README): headings, fences, a table rule,
 * or a list of several lines — what LaTeX input also takes as Markdown (Typora parses Markdown pastes in every mode).
 * LaTeX source (\documentclass, \section{…}) is not.
 */
export function looksLikeMarkdownDocument(text: string): boolean {
  if (/\\(documentclass|section\*?\{|subsection\*?\{|begin\{document\})/.test(text)) return false;
  const lines = text.split(/\r?\n/);
  const listLines = lines.filter((line) => /^\s*([-*+]|\d{1,3}[.)])\s+\S/.test(line)).length;
  return (
    lines.some((line) => /^#{1,6}\s+\S/.test(line)) ||
    lines.some((line) => /^\s*(```|~~~)/.test(line)) ||
    lines.some((line) => /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/.test(line)) ||
    listLines >= 2 ||
    (listLines >= 1 && /\*\*[^*\n]+\*\*/.test(text))
  );
}

// ---------------------------------------------------------------- writing --

/** Footnote numbers (by node) and reference list numbers (by entry id). */
interface NoteNumbers {
  footnotes: Map<Node, number>;
  entries: Map<string, number>;
}

function noteNumbers(doc: Node): NoteNumbers {
  const footnotes = new Map<Node, number>();
  doc.descendants((node) => {
    if (node.type.name === "footnote") footnotes.set(node, footnotes.size + 1);
    return true;
  });
  return { footnotes, entries: new Map(bibliographyEntries(doc).map((entry) => [entry.id, entry.number])) };
}

/** The footnotes' texts as Markdown footnote definitions ([^1]: …), after the document. */
function footnoteDefinitions(doc: Node): string {
  const lines: string[] = [];
  doc.descendants((node) => {
    if (node.type.name === "footnote") lines.push(`[^${lines.length + 1}]: ${String(node.attrs.text).replace(/\n+/g, " ")}`);
    return true;
  });
  return lines.join("\n");
}

export interface MarkdownExportOptions {
  /** The reference list's title (the editor shows it above the list); none: no title line. */
  bibliographyTitle?: string;
  /** A formula's LaTeX as it leaves the app (the same conversions as Copy LaTeX). */
  formulaLatex?: (latex: string) => string;
  /**
   * References to numbered formulas: `$\eqref{label}$` with a `\label` in the formula, which Typora and MathJax turn
   * into links (KaTeX viewers such as VS Code show `\label` as an error), or the number as plain text "(n)".
   */
  references?: "eqref" | "plain";
  /** The Markdown file the document was read from: blocks still as they were are written back unchanged. */
  source?: MarkdownSource | null;
  /** Where an image's file is, relative to the .md (default: assets/fig-n.ext in the order of appearance). */
  imagePath?: (key: string) => string | null;
}

const htmlAttr = (value: string) => value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
const percent = (width: number) => `${Math.round(width * 1000) / 10}%`;

/** A formula id as a LaTeX label: as read from a file, or eq:<id> for formulas made here. */
function labelOf(id: string, source: MarkdownSource | null | undefined): string {
  return source?.labels.has(id) || id.startsWith("eq:") ? id : `eq:${id}`;
}

/** The ids of the formulas the document refers to. */
function referencedIds(doc: Node): Set<string> {
  const ids = new Set<string>();
  doc.descendants((node) => {
    if (node.type.name === "eqref" && node.attrs.target) ids.add(node.attrs.target);
    return true;
  });
  return ids;
}

/** The numbers of the numbered display formulas: by node (in document order) and by id (for references). */
function numbering(doc: Node): { byNode: Map<Node, string>; byId: Map<string, string>; brackets: "()" | "[]" } {
  // the document's numbering format (settings.ts): 3, 2.3 or 2-3
  const labels = equationLabels(doc);
  const byNode = new Map<Node, string>();
  const byId = new Map<string, string>();
  let n = 0;
  doc.descendants((node) => {
    if (node.type.name === "math_display" && node.attrs.numbered) {
      n += 1;
      const label = (node.attrs.id && labels.get(node.attrs.id)) || String(n);
      byNode.set(node, label);
      if (node.attrs.id) byId.set(node.attrs.id, label);
    }
    return true;
  });
  return { byNode, byId, brackets: documentSettings(doc).equationBrackets };
}

/** Whether the editor starts this top-level block below a wrapped image (document.css: clear). */
function clearsWrap(node: Node): boolean {
  if (node.type.name !== "paragraph") return node.isBlock;
  return isImageLine(node) || (isImage(node.firstChild) && !!node.firstChild!.attrs.wrap);
}

/** Whether a wrapped image above the block at `index` may still float beside it (only body paragraphs between). */
function wrapGoesOn(doc: Node, index: number): boolean {
  for (let i = index - 1; i >= 0; i--) {
    const block = doc.child(i);
    if (block.type.name !== "paragraph" || isImageLine(block)) return false;
    if (isImage(block.firstChild) && block.firstChild!.attrs.wrap) return true;
  }
  return false;
}

function serializerFor(
  numbers: { byNode: Map<Node, string>; byId: Map<string, string>; brackets: "()" | "[]" },
  options: MarkdownExportOptions,
  referenced: Set<string> = new Set(),
  notes: NoteNumbers = { footnotes: new Map(), entries: new Map() },
): MarkdownSerializer {
  const live = (options.references ?? "eqref") === "eqref";
  // New and changed blocks follow the file's own habits (Typora keeps a list's marker the same way).
  const style = options.source?.style ?? { bullet: "-", delimiter: ".", setext: false };
  const latexOf = (latex: string) => (options.formulaLatex ?? ((l: string) => l))(latex).trim();
  const pathOf = (key: string) => (options.imagePath ?? (() => null))(key) ?? key;
  /** An image as the HTML GitHub, Typora, Obsidian and VS Code show at its width (DESIGN.md §11.1). */
  const imageHtml = (node: Node) =>
    `<img src="${htmlAttr(pathOf(node.attrs.src))}" alt="${htmlAttr(node.attrs.alt || "")}" width="${percent(node.attrs.width)}"${node.attrs.wrap ? ` align="${node.attrs.wrap}"` : ""}>`;
  const nodes: ConstructorParameters<typeof MarkdownSerializer>[0] = {
    paragraph(state, node) {
      if (isImageLine(node)) {
        // an image line: one <p> (align on <img> does not centre), the images a space apart
        const images: string[] = [];
        node.forEach((child) => {
          if (isImage(child)) images.push(imageHtml(child));
        });
        // the gap the editor leaves (0.02 of the line, about 1em in a viewer's column): an em space
        state.write(`<p align="${(node.attrs.align as string | null) ?? "left"}">${images.join("&emsp;")}</p>`);
        state.closeBlock(node);
        return;
      }
      const align = node.attrs.align as string | null;
      if (align === "center" || align === "right") {
        // Markdown has no alignment: HTML, as Typora writes it.
        state.write(`<p align="${align}">`);
        state.renderInline(node);
        state.write("</p>");
      } else state.renderInline(node);
      state.closeBlock(node);
    },
    title(state, node) {
      state.write(MARKDOWN_TITLE_OPEN);
      state.renderInline(node, false);
      state.write(MARKDOWN_TITLE_CLOSE);
      state.closeBlock(node);
    },
    heading(state, node) {
      // A file written with underlined (setext) headings gets its new level-1/2 headings the same way.
      if (style.setext && node.attrs.level <= 2) {
        state.renderInline(node, false);
        state.write(`\n${(node.attrs.level === 1 ? "=" : "-").repeat(Math.max(3, Math.min(60, node.textContent.length * 2)))}`);
        state.closeBlock(node);
        return;
      }
      state.write(`${"#".repeat(node.attrs.level)} `);
      state.renderInline(node, false);
      state.closeBlock(node);
    },
    blockquote(state, node) {
      state.wrapBlock("> ", null, node, () => {
        // an alert: its kind on the first line (GitHub's form, which Typora reads)
        if (node.attrs.alert) {
          state.write(`[!${String(node.attrs.alert).toUpperCase()}]`);
          state.ensureNewLine();
        }
        state.renderContent(node);
      });
    },
    image_board(state, node) {
      // a plain row as an image line; otherwise rows of images with their margins (Typora, Obsidian, VS Code show
      // them exactly; GitHub keeps the rows), or, when images overlap, one group with each image placed; every image
      // carries its place (data-x, data-top) for VisualTeX
      const row = canonicalRow(node);
      if (row) {
        state.write(`<p align="${row.align}">${row.items.map((item) => imageHtml(item.node)).join("&emsp;")}</p>`);
        state.closeBlock(node);
        return;
      }
      const pct = (value: number) => `${Math.round(value * 10000) / 100}%`;
      const f3 = (value: number) => String(Math.round(value * 1000) / 1000);
      const img = (item: ReturnType<typeof boardItems>[number], style: string) =>
        `<img src="${htmlAttr(pathOf(item.node.attrs.src))}" alt="${htmlAttr(item.node.attrs.alt || "")}" width="${percent(item.width)}" data-x="${f3(item.x)}" data-top="${f3(item.top)}" style="${style}">`;
      const rows = boardRows(node);
      let html: string;
      if (rows) {
        let previousBottom = 0;
        html = `<div data-vtx-board="1">${rows
          .map((items) => {
            const rowTop = Math.min(...items.map((item) => item.top));
            const rowBottom = Math.max(...items.map((item) => item.top + item.height));
            const marginTop = rowTop - previousBottom;
            previousBottom = rowBottom;
            let right = 0;
            const images = items
              .map((item) => {
                const gap = item.x - right;
                right = item.x + item.width;
                return img(item, `margin:0 0 ${pct(rowBottom - item.top - item.height)} ${pct(gap)};vertical-align:baseline`);
              })
              .join("");
            return `<p style="margin:${pct(marginTop)} 0 0 0;line-height:0">${images}</p>`;
          })
          .join("")}</div>`;
      } else {
        const height = boardHeight(node);
        html = `<div data-vtx-board="1" style="position:relative;height:0;padding-bottom:${pct(height)}">${boardItems(node)
          .map((item) => img(item, `position:absolute;left:${pct(item.x)};top:${pct(item.top / height)};width:${pct(item.width)}`))
          .join("")}</div>`;
      }
      state.write(html);
      state.closeBlock(node);
    },
    board_image() {},
    toc(state, node) {
      state.write("[toc]");
      state.closeBlock(node);
    },
    html_inline(state, node) {
      state.write(String(node.attrs.html ?? ""));
    },
    html_block(state, node) {
      state.text(String(node.attrs.html ?? ""), false);
      state.closeBlock(node);
    },
    bullet_list(state, node) {
      state.renderList(node, "  ", () => `${style.bullet} `);
    },
    ordered_list(state, node) {
      const start = (node.attrs.start as number) || 1;
      const width = String(start + node.childCount - 1).length;
      state.renderList(node, " ".repeat(width + 2), (i) => {
        const n = String(start + i);
        return `${" ".repeat(width - n.length)}${n}${style.delimiter} `;
      });
    },
    list_item(state, node) {
      if (node.attrs.checked !== null) state.write(node.attrs.checked ? "[x] " : "[ ] ");
      state.renderContent(node);
    },
    code_block(state, node) {
      const text = node.textContent;
      const ticks = "`".repeat(Math.max(3, ...(text.match(/`{3,}/g) ?? []).map((run) => run.length + 1)));
      state.write(`${ticks}${node.attrs.language ?? ""}\n`);
      state.text(text, false);
      state.ensureNewLine();
      state.write(ticks);
      state.closeBlock(node);
    },
    page_command(state, node) {
      // a page break as Typora and the browsers print one; the other commands kept as a comment
      if (node.attrs.command === "newpage" || node.attrs.command === "clearpage") state.write('<div style="page-break-after: always;"></div>');
      else state.write(`<!-- ${pageCommandLatex(node.attrs.command, node.attrs.value)} -->`);
      state.closeBlock(node);
    },
    horizontal_rule(state, node) {
      state.write("---");
      state.closeBlock(node);
    },
    math_display(state, node) {
      const number = numbers.byNode.get(node);
      // A formula the text refers to carries its label, so the references can point at it.
      const label = live && number !== undefined && node.attrs.id && referenced.has(node.attrs.id) ? ` \\label{${labelOf(node.attrs.id, options.source)}}` : "";
      state.write("$$");
      state.ensureNewLine();
      const tag = number === undefined ? "" : numbers.brackets === "[]" ? ` \\tag*{[${number}]}` : ` \\tag{${number}}`;
      state.text(latexOf(node.attrs.latex) + tag + (tag ? label : ""), false);
      state.ensureNewLine();
      state.write("$$");
      state.closeBlock(node);
    },
    math_inline(state, node) {
      state.write(`$${latexOf(node.attrs.latex)}$`);
    },
    image(state, node) {
      // a wrapped image opens its paragraph's line, the text beside it (an <img> alone on the line above would make
      // the paragraph raw HTML)
      state.write(imageHtml(node) + (node.attrs.wrap ? " " : ""));
    },
    eqref(state, node) {
      const number = node.attrs.target ? numbers.byId.get(node.attrs.target) : undefined;
      if (number === undefined) state.write("(??)");
      else state.write(live ? `$\\eqref{${labelOf(node.attrs.target, options.source)}}$` : numbers.brackets === "[]" ? `[${number}]` : `(${number})`);
    },
    footnote(state, node) {
      state.write(`[^${notes.footnotes.get(node) ?? "?"}]`);
    },
    cite(state, node) {
      state.write(citeLabel((node.attrs.targets as string[]).map((id) => notes.entries.get(id) ?? null)));
    },
    bibliography(state, node) {
      if (options.bibliographyTitle) {
        state.write(`**${options.bibliographyTitle}**`);
        state.closeBlock(node);
      }
      node.forEach((entry, _offset, index) => {
        state.write(`[${index + 1}] `);
        state.renderInline(entry);
        state.closeBlock(entry);
      });
    },
    bib_entry(state, node) {
      state.renderInline(node);
      state.closeBlock(node);
    },
    hard_break(state, node, parent, index) {
      for (let i = index + 1; i < parent.childCount; i++) {
        if (parent.child(i).type !== node.type) {
          state.write("\\\n");
          return;
        }
      }
    },
    text(state, node) {
      state.text(node.text ?? "");
    },
    table(state, node) {
      // GFM: the first row is the header row; a cell is one line (its paragraphs joined by <br>, formulas inline).
      const rows: string[][] = [];
      node.forEach((row) => {
        const cells: string[] = [];
        row.forEach((cell) => {
          const parts: string[] = [];
          cell.forEach((block) => {
            if (block.type.name === "math_display") parts.push(`$${latexOf(block.attrs.latex)}$`);
            else parts.push(serializer.serialize(block.type.schema.topNodeType.create(null, block)).trim().replace(/\n/g, " "));
          });
          cells.push(parts.join("<br>").replace(/\|/g, "\\|") || " ");
        });
        rows.push(cells);
      });
      const width = Math.max(...rows.map((r) => r.length));
      const line = (cells: string[]) => `| ${Array.from({ length: width }, (_, i) => cells[i] ?? " ").join(" | ")} |`;
      // the columns' alignment: the header row's cells'
      const aligns: Array<string | null> = [];
      node.firstChild?.forEach((cell) => aligns.push(cell.attrs.align ?? null));
      const rule = (i: number) => ({ left: " :--- ", center: " :---: ", right: " ---: " })[aligns[i] as "left"] ?? " --- ";
      const out = [line(rows[0] ?? []), `|${Array.from({ length: width }, (_, i) => rule(i)).join("|")}|`, ...rows.slice(1).map(line)];
      state.write(out.join("\n"));
      state.closeBlock(node);
    },
    table_row() {},
    table_cell() {},
    table_header() {},
  };
  const marks: ConstructorParameters<typeof MarkdownSerializer>[1] = {
    link: {
      open: "[",
      close: (_state, mark) => `](${String(mark.attrs.href).replace(/[()]/g, (c) => `\\${c}`)}${mark.attrs.title ? ` "${String(mark.attrs.title).replace(/"/g, '\\"')}"` : ""})`,
    },
    strong: emphasis("**", "strong"),
    em: emphasis("*", "em"),
    underline: { open: "<u>", close: "</u>", mixable: true },
    // as HTML, which Typora, GitHub and Obsidian all show (x^2^ needs Typora's option)
    superscript: { open: "<sup>", close: "</sup>", mixable: true },
    subscript: { open: "<sub>", close: "</sub>", mixable: true },
    kbd: { open: "<kbd>", close: "</kbd>" },
    strike: { open: "~~", close: "~~", mixable: true, expelEnclosingWhitespace: true },
    code: {
      open: (_state, _mark, parent, index) => backticks(parent.child(index).text ?? "", -1),
      close: (_state, _mark, parent, index) => backticks(parent.child(index - 1).text ?? "", 1),
      escape: false,
    },
    // Colour and size as the HTML Typora keeps and shows (GitHub shows the text plain).
    textColor: { open: (_state, mark) => `<span style="color:${mark.attrs.color}">`, close: "</span>" },
    highlight: {
      open: (_state, mark) => (mark.attrs.color === MARKDOWN_HIGHLIGHT ? "==" : `<span style="background-color:${mark.attrs.color}">`),
      close: (_state, mark) => (mark.attrs.color === MARKDOWN_HIGHLIGHT ? "==" : "</span>"),
    },
    fontSize: {
      open: (_state, mark) => `<span style="font-size:${FONT_SIZES.find((size) => size.name === mark.attrs.size)?.ratio ?? 1}em">`,
      close: "</span>",
    },
  };
  // A wrapped image floats beside the paragraphs after it, as in the editor; what the editor starts below it (a
  // heading, a list, an image line, the next wrapped paragraph…) gets <br clear="all"> first (GitHub keeps it).
  for (const [name, write] of Object.entries(nodes)) {
    nodes[name] = (state, node, parent, index) => {
      if (parent?.type.name === "doc" && clearsWrap(node) && wrapGoesOn(parent, index)) {
        state.write('<br clear="all">');
        state.closeBlock(node);
      }
      write(state, node, parent, index);
    };
  }
  // Besides CommonMark's characters: "$" (math), "==" (highlight) and "<" before a tag name (HTML), all literal here.
  // Also "^" before a second one (x^2^ is superscript when read) and ":" of an emoji short code.
  const serializer = new MarkdownSerializer(nodes, marks, { escapeExtraCharacters: /[$]|==|<(?=[A-Za-z\/!?])|\^(?=[^\s^]+\^)|:(?=[a-z_+][a-z0-9_+-]*:)/g });
  return serializer;
}

const PUNCTUATION = /[\p{P}\p{S}]/u;
const SPACE = /\s/u;

/** The text a run of inline nodes writes at one end (a formula writes "$", a reference "(", a line break "\n"). */
function edgeChar(node: Node | null, side: -1 | 1): string {
  if (!node) return "";
  if (node.isText) return side < 0 ? node.text!.slice(-1) : node.text!.slice(0, 1);
  if (node.type.name === "hard_break") return "\n";
  if (node.type.name === "eqref") return side < 0 ? ")" : "(";
  return "$";
}

/**
 * Bold and italic as `**` / `*`, unless CommonMark could not read the delimiters back: next to punctuation inside and
 * a letter outside (`**中文：**文字`, CommonMark #650), where the run is written as <strong> / <em> instead.
 */
function emphasis(delimiter: string, tag: string): ConstructorParameters<typeof MarkdownSerializer>[1][string] {
  const run = (mark: Mark, parent: Node, index: number, open: boolean): { first: number; end: number } => {
    let first = open ? index : index - 1;
    let end = open ? index : index;
    if (open) {
      end = index;
      while (end < parent.childCount && mark.isInSet(parent.child(end).marks)) end++;
    } else {
      while (first > 0 && mark.isInSet(parent.child(first - 1).marks)) first--;
    }
    return { first, end };
  };
  const asHtml = (mark: Mark, parent: Node, index: number, open: boolean) => {
    const { first, end } = run(mark, parent, index, open);
    if (first >= end) return false;
    const inside = (side: -1 | 1) => {
      // Whitespace at the run's ends goes outside the delimiters (expelEnclosingWhitespace).
      for (let i = side < 0 ? first : end - 1; side < 0 ? i < end : i >= first; i += side < 0 ? 1 : -1) {
        const node = parent.child(i);
        const text = node.isText ? node.text! : edgeChar(node, side);
        const trimmed = side < 0 ? text.replace(/^\s+/, "") : text.replace(/\s+$/, "");
        if (trimmed) return { char: side < 0 ? trimmed[0] : trimmed.slice(-1), padded: trimmed.length !== text.length };
      }
      return null;
    };
    const head = inside(-1);
    const tail = inside(1);
    if (!head || !tail) return false;
    const before = head.padded ? " " : edgeChar(first > 0 ? parent.child(first - 1) : null, -1);
    const after = tail.padded ? " " : edgeChar(end < parent.childCount ? parent.child(end) : null, 1);
    const opens = !PUNCTUATION.test(head.char) || !before || SPACE.test(before) || PUNCTUATION.test(before);
    const closes = !PUNCTUATION.test(tail.char) || !after || SPACE.test(after) || PUNCTUATION.test(after);
    return !(opens && closes);
  };
  return {
    open: (_state, mark, parent, index) => (asHtml(mark, parent, index, true) ? `<${tag}>` : delimiter),
    close: (_state, mark, parent, index) => (asHtml(mark, parent, index, false) ? `</${tag}>` : delimiter),
    mixable: true,
    expelEnclosingWhitespace: true,
  };
}

/** Backticks for inline code: one more than the longest run inside, with a space when it starts or ends with one. */
function backticks(text: string, side: -1 | 1): string {
  const longest = Math.max(0, ...(text.match(/`+/g) ?? []).map((run) => run.length));
  const ticks = "`".repeat(longest + 1);
  const pad = longest > 0 && (side < 0 ? text.startsWith("`") : text.endsWith("`")) ? " " : "";
  return side < 0 ? ticks + pad : pad + ticks;
}

/** The document as Markdown. */
export function documentToMarkdown(doc: Node, options: MarkdownExportOptions = {}): string {
  if (!options.imagePath) {
    const paths = imageFiles(doc, "assets").paths;
    options = { ...options, imagePath: (key) => paths.get(key) ?? null };
  }
  const numbers = numbering(doc);
  const notes = noteNumbers(doc);
  const serializer = serializerFor(numbers, options, referencedIds(doc), notes);
  const source = options.source;
  // the footnotes' texts follow the document, as Markdown footnote definitions
  const definitions = footnoteDefinitions(doc);
  if (!source) return `${serializer.serialize(doc, { tightLists: true }).trim()}\n${definitions ? `\n${definitions}\n` : ""}`;
  const text = withSource(doc, source, numbers, (node, index) => {
    if (index === 0 && source.frontMatter && node.type.name === "code_block" && node.attrs.language === "yaml") return `---\n${node.textContent}\n---`;
    return serializer.serialize(doc.type.create(null, [node]), { tightLists: true }).replace(/^\n+|\n+$/g, "");
  });
  if (!definitions) return text;
  const newline = source.crlf ? "\r\n" : "\n";
  return `${text.replace(/\s+$/, "")}${newline}${newline}${definitions.replace(/\n/g, newline)}${newline}`;
}

// ------------------------------------------------------- keeping the file --

/** A Markdown file as it was read: what each top-level block was, and where its text is. */
export interface MarkdownSource {
  /** the file, newlines as \n */
  text: string;
  /** the file had \r\n line ends */
  crlf: boolean;
  blocks: Array<{ signature: string; start: number; end: number }>;
  /** the first block is YAML front matter (a yaml code block in the document) */
  frontMatter: boolean;
  /** formula ids that were \label{…}s in the file: written back as they were */
  labels: Set<string>;
  /** the file's habits: its bullet, its ordered-list delimiter, underlined headings */
  style: { bullet: "-" | "*" | "+"; delimiter: "." | ")"; setext: boolean };
}

/** The markers a file uses most: bullets, ordered-list delimiters, underlined or # headings (code fences skipped). */
function styleOf(text: string): MarkdownSource["style"] {
  const lines = text.split("\n");
  const bullets: Record<string, number> = { "-": 0, "*": 0, "+": 0 };
  const delimiters: Record<string, number> = { ".": 0, ")": 0 };
  let setext = 0;
  let atx = 0;
  let fence = false;
  lines.forEach((line, i) => {
    if (/^\s*(```|~~~)/.test(line)) fence = !fence;
    if (fence) return;
    const bullet = /^\s*([-*+])\s+\S/.exec(line);
    if (bullet && !/^\s*([-*_])(\s*\1){2,}\s*$/.test(line)) bullets[bullet[1]]++;
    const ordered = /^\s*\d{1,9}([.)])\s+\S/.exec(line);
    if (ordered) delimiters[ordered[1]]++;
    if (/^#{1,2}\s/.test(line)) atx++;
    if (i > 0 && /^(=+|-+)\s*$/.test(line) && /\S/.test(lines[i - 1]) && !/^\s*([-*+]|\d+[.)])\s/.test(lines[i - 1])) setext++;
  });
  const most = <K extends string>(counts: Record<K, number>, fallback: K): K =>
    (Object.keys(counts) as K[]).reduce((best, key) => (counts[key] > counts[best] ? key : best), fallback);
  return { bullet: most(bullets, "-") as "-" | "*" | "+", delimiter: most(delimiters, ".") as "." | ")", setext: setext > atx };
}

/** What a block is, for "is it still as it was read": its content without ids, and its formula numbers. */
function signature(node: Node, numbers: { byNode: Map<Node, string>; byId: Map<string, string>; brackets: "()" | "[]" }): string {
  const marks: string[] = [];
  const visit = (n: Node) => {
    if (n.type.name === "math_display" && n.attrs.numbered) marks.push(`#${numbers.byNode.get(n) ?? "?"}`);
    if (n.type.name === "eqref") marks.push(`r${n.attrs.target ? (numbers.byId.get(n.attrs.target) ?? "?") : "?"}`);
  };
  visit(node);
  node.descendants((n) => {
    visit(n);
    return true;
  });
  return JSON.stringify(node.toJSON(), (key, value) => (key === "id" ? undefined : value)) + marks.join(",");
}

/**
 * The document as Markdown with the blocks that are still as they were read written back character for character,
 * the space between two such neighbours too; only new or changed blocks are written anew (Typora keeps a file
 * the same way, typora-issues-report.md). An unchanged document gives the file back exactly.
 */
function withSource(
  doc: Node,
  source: MarkdownSource,
  numbers: { byNode: Map<Node, string>; byId: Map<string, string>; brackets: "()" | "[]" },
  write: (node: Node, index: number) => string,
): string {
  const parts: Array<{ text: string; original: number }> = [];
  let next = 0;
  doc.forEach((node, _offset, index) => {
    const sig = signature(node, numbers);
    let original = -1;
    for (let i = next; i < source.blocks.length; i++) {
      if (source.blocks[i].signature === sig) {
        original = i;
        break;
      }
    }
    if (original >= 0) {
      next = original + 1;
      const block = source.blocks[original];
      parts.push({ text: source.text.slice(block.start, block.end), original });
    } else parts.push({ text: write(node, index), original: -1 });
  });
  // Blocks edited in place (as many new blocks between two kept ones as were there in the file) take the places of
  // the ones they replace, so the space around them stays as it was too.
  const place = parts.map((part) => part.original);
  for (let i = 0; i < parts.length; ) {
    if (place[i] >= 0) {
      i++;
      continue;
    }
    let j = i;
    while (j < parts.length && place[j] < 0) j++;
    const before = i > 0 ? place[i - 1] : -1;
    const after = j < parts.length ? place[j] : source.blocks.length;
    if (after - before - 1 === j - i) for (let k = i; k < j; k++) place[k] = before + 1 + (k - i);
    i = j;
  }
  const gapBefore = (i: number) => source.text.slice(i > 0 ? source.blocks[i - 1].end : 0, source.blocks[i].start);
  let out = "";
  parts.forEach((part, i) => {
    if (i === 0) out += place[0] === 0 ? gapBefore(0) : "";
    else if (place[i] > 0 && place[i - 1] === place[i] - 1) out += gapBefore(place[i]);
    else {
      // Between rewritten neighbours a blank line; what stood before the block in the file (link definitions) stays.
      const kept = place[i] > 0 ? gapBefore(place[i]).trim() : "";
      out += kept ? `\n\n${kept}\n\n` : "\n\n";
    }
    out += part.text;
  });
  const tail = source.blocks.length ? source.text.slice(source.blocks[source.blocks.length - 1].end) : source.text;
  if (parts.length && place[parts.length - 1] === source.blocks.length - 1) out += tail;
  else out += tail.trim() ? `\n\n${tail.trim()}\n` : "\n";
  return source.crlf ? out.replace(/\n/g, "\r\n") : out;
}

/** A Markdown file as a document, with what is needed to write its unchanged blocks back as they were. */
export function readMarkdown(schema: Schema, raw: string): { doc: Node; source: MarkdownSource | null } {
  const text = raw.replace(/\r\n?/g, "\n");
  // YAML front matter (Typora, Obsidian, Hugo): a yaml code block, written back as front matter.
  const front = /^---[ \t]*\n([\s\S]*?)\n(?:---|\.\.\.)[ \t]*(?:\n|$)/.exec(text);
  const bodyStart = front ? front[0].length : 0;
  const body = text.slice(bodyStart);
  const parser = parserFor(schema);
  let doc = linkReferences(resolveLatexMarkers(parser.parse(body)));
  if (front) {
    const yaml = schema.nodes.code_block.create({ language: "yaml" }, front[1] ? schema.text(front[1]) : null);
    doc = doc.type.create(doc.attrs, [yaml, ...childrenOf(doc)]);
  }
  // The top-level blocks' lines in the file.
  const lineStarts = [0];
  for (let i = 0; i < body.length; i++) if (body[i] === "\n") lineStarts.push(i + 1);
  const offset = (line: number) => bodyStart + (line < lineStarts.length ? lineStarts[line] : body.length);
  const ranges: Array<{ start: number; end: number }> = [];
  if (front) ranges.push({ start: 0, end: front[0].replace(/\n$/, "").length });
  for (const token of parser.tokenizer.parse(body, {})) {
    if (token.level !== 0 || token.nesting < 0 || !token.map) continue;
    const start = offset(token.map[0]);
    let end = offset(token.map[1]);
    while (end > start && text[end - 1] === "\n") end--;
    ranges.push({ start, end });
  }
  const labels = new Set<string>();
  doc.descendants((node) => {
    if (node.type.name === "math_display" && node.attrs.id) labels.add(node.attrs.id);
    return true;
  });
  // Blocks and lines must pair up one to one; otherwise the file is simply written anew.
  if (ranges.length !== doc.childCount) return { doc, source: null };
  const numbers = numbering(doc);
  const blocks = childrenOf(doc).map((node, i) => ({ signature: signature(node, numbers), ...ranges[i] }));
  return { doc, source: { text, crlf: /\r\n/.test(raw), blocks, frontMatter: !!front, labels, style: styleOf(body) } };
}

/** The Markdown file the document was read from (set by the transaction that loads it; a new document drops it). */
export const markdownSourceKey = new PluginKey<MarkdownSource | null>("visualtex-document-markdown-source");

export function markdownSourcePlugin(): Plugin<MarkdownSource | null> {
  return new Plugin<MarkdownSource | null>({
    key: markdownSourceKey,
    state: {
      init: () => null,
      apply: (tr, value) => {
        const meta = tr.getMeta(markdownSourceKey) as MarkdownSource | null | undefined;
        return meta !== undefined ? meta : value;
      },
    },
  });
}

export function markdownSourceOf(state: EditorState): MarkdownSource | null {
  return markdownSourceKey.getState(state) ?? null;
}

function childrenOf(node: Node): Node[] {
  const children: Node[] = [];
  node.forEach((child) => children.push(child));
  return children;
}

/** A selection's content as Markdown (Copy as Markdown). */
export function sliceToMarkdown(slice: Slice, schema: Schema, options: MarkdownExportOptions = {}): string {
  const doc = schema.topNodeType.createAndFill(null, slice.content);
  if (!doc) {
    // Inline content: in a paragraph.
    const paragraph = schema.nodes.paragraph.create(null, slice.content);
    return serializerFor(numbering(schema.topNodeType.create(null, paragraph)), options).serialize(schema.topNodeType.create(null, paragraph)).trim();
  }
  return serializerFor(numbering(doc), options).serialize(doc, { tightLists: true }).trim();
}
