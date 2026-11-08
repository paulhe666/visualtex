// SyncTeX, read here (one reader for the app and the dev bridge): a page point → the source line that made it,
// a source line → where it is on the pages. Positions are in bp from the page's top left corner.

export interface SyncRecord {
  /** index into `files` */
  file: number;
  line: number;
  /** "(" hbox, "[" vbox, "h"/"v" void boxes, "x" "k" "g" "$" points */
  kind: string;
  h: number;
  v: number;
  width: number;
  height: number;
  depth: number;
  /** the enclosing box (index into the page's records), -1 at the top */
  parent: number;
}

export interface SyncPage {
  records: SyncRecord[];
}

const SP_PER_BP = 65781.76;

export class SyncTex {
  /** input files, by synctex tag (paths as TeX recorded them) */
  readonly files: string[] = [];
  readonly pages = new Map<number, SyncPage>();

  static parse(text: string): SyncTex {
    const sync = new SyncTex();
    const tags = new Map<number, number>();
    let unit = 1;
    let xOffset = 0;
    let yOffset = 0;
    let page: SyncPage | null = null;
    const stack: number[] = [];
    const lines = text.split("\n");
    const toBp = (value: string, offset: number) => (Number(value) * unit + offset) / SP_PER_BP;
    for (const raw of lines) {
      const line = raw.replace(/\r$/, "");
      if (!line) continue;
      const c = line[0];
      if (line.startsWith("Input:")) {
        const m = /^Input:(\d+):(.*)$/.exec(line);
        if (m) {
          tags.set(Number(m[1]), sync.files.length);
          sync.files.push(m[2]);
        }
        continue;
      }
      if (!page) {
        if (line.startsWith("Unit:")) unit = Number(line.slice(5)) || 1;
        else if (line.startsWith("X Offset:")) xOffset = Number(line.slice(9)) || 0;
        else if (line.startsWith("Y Offset:")) yOffset = Number(line.slice(9)) || 0;
      }
      if (c === "{") {
        page = { records: [] };
        sync.pages.set(Number(line.slice(1)), page);
        stack.length = 0;
        continue;
      }
      if (c === "}") {
        page = null;
        continue;
      }
      if (!page) continue;
      if (c === "]" || c === ")") {
        stack.pop();
        continue;
      }
      if (!"[(hvxkg$".includes(c)) continue;
      // kind tag,line:h,v[:W[,H,D]]
      const m = /^.(\d+),(\d+)(?:,\d+)?:(-?\d+),(-?\d+)(?::(-?\d+)(?:,(-?\d+),(-?\d+))?)?/.exec(line);
      if (!m) continue;
      const record: SyncRecord = {
        file: tags.get(Number(m[1])) ?? -1,
        line: Number(m[2]),
        kind: c,
        h: toBp(m[3], xOffset),
        v: toBp(m[4], yOffset),
        width: m[5] ? (Number(m[5]) * unit) / SP_PER_BP : 0,
        height: m[6] ? (Number(m[6]) * unit) / SP_PER_BP : 0,
        depth: m[7] ? (Number(m[7]) * unit) / SP_PER_BP : 0,
        parent: stack.length ? stack[stack.length - 1] : -1,
      };
      page.records.push(record);
      if (c === "[" || c === "(") stack.push(page.records.length - 1);
    }
    return sync;
  }

  /**
   * The source line at a page point: the smallest box around the point, the record in it nearest the point
   * (its line), preferring files `wanted` accepts (the project's own over classes and packages).
   */
  edit(pageNumber: number, x: number, y: number, wanted: (file: number) => boolean = () => true): { file: number; line: number } | null {
    const page = this.pages.get(pageNumber);
    if (!page) return null;
    const recs = page.records;
    let box = -1;
    let boxArea = Infinity;
    recs.forEach((r, i) => {
      if (r.kind !== "(" && r.kind !== "[") return;
      const top = r.v - r.height;
      const bottom = r.v + Math.max(r.depth, 0);
      if (x >= r.h && x <= r.h + r.width && y >= top - 1 && y <= bottom + 1) {
        const area = Math.max(r.width, 0.1) * Math.max(bottom - top, 0.1);
        if (area < boxArea && (r.kind === "(" || box < 0)) {
          box = i;
          boxArea = area;
        }
      }
    });
    const distance = (r: SyncRecord) => {
      const top = r.kind === "(" || r.kind === "[" || r.kind === "h" ? r.v - r.height : r.v - 8;
      const bottom = r.kind === "(" || r.kind === "[" || r.kind === "h" ? r.v + r.depth : r.v + 2;
      const dy = y < top ? top - y : y > bottom ? y - bottom : 0;
      const right = r.h + Math.max(r.width, 0);
      const dx = x < r.h ? r.h - x : x > right ? (r.width > 0 ? x - right : (x - r.h) * 0.25) : 0;
      return dy * 4 + dx;
    };
    const pick = (list: SyncRecord[]) => {
      let best: SyncRecord | null = null;
      let bestD = Infinity;
      for (const r of list) {
        if (r.line <= 0 || r.file < 0 || !wanted(r.file)) continue;
        const d = distance(r);
        if (d < bestD) {
          best = r;
          bestD = d;
        }
      }
      return best;
    };
    // inside the box, its glyph-level records first: a line box carries the line where its paragraph ended
    const points = (list: SyncRecord[]) => list.filter((r) => r.kind === "x" || r.kind === "g" || r.kind === "k" || r.kind === "$");
    const inBox = box >= 0 ? recs.filter((r, i) => i === box || r.parent === box) : [];
    const found = pick(points(inBox)) ?? pick(inBox) ?? pick(points(recs)) ?? pick(recs);
    return found ? { file: found.file, line: found.line } : null;
  }

  /**
   * Where a source line is: the first page with records of the line (or the nearest line before it that has
   * some), and the area they cover.
   */
  view(file: number, line: number): { page: number; x: number; y: number; width: number; height: number } | null {
    for (let delta = 0; delta < 40; delta++) {
      for (const target of delta === 0 ? [line] : [line - delta, line + delta]) {
        if (target <= 0) continue;
        const found = this.place(file, target);
        if (found) return found;
      }
    }
    return null;
  }

  /**
   * The records of one line, grouped into bands of a page by height; the band with the most of them (the line's
   * text, not a running head the page break made while TeX read the line) is the place.
   */
  private place(file: number, line: number): { page: number; x: number; y: number; width: number; height: number } | null {
    // glyph-level records (a line box carries the line where its paragraph was broken); boxes when there are none
    return this.band(file, line, (r) => r.kind === "x" || r.kind === "g" || r.kind === "k" || r.kind === "$") ?? this.band(file, line, (r) => (r.kind === "(" && r.width > 0) || r.kind === "h");
  }

  private band(file: number, line: number, kinds: (r: SyncRecord) => boolean): { page: number; x: number; y: number; width: number; height: number } | null {
    let best: { page: number; count: number; recs: SyncRecord[] } | null = null;
    for (const p of [...this.pages.keys()].sort((a, b) => a - b)) {
      const recs = this.pages.get(p)!.records.filter((r) => r.file === file && r.line === line && kinds(r));
      if (!recs.length) continue;
      const bands: SyncRecord[][] = [];
      for (const r of [...recs].sort((a, b) => a.v - b.v)) {
        const band = bands[bands.length - 1];
        if (band && r.v - band[band.length - 1].v < 16) band.push(r);
        else bands.push([r]);
      }
      for (let band of bands) {
        // a band taller than a few lines (text a macro made, all tagged with this line): its fullest baseline
        if (band[band.length - 1].v - band[0].v > 60) {
          const lines = new Map<number, SyncRecord[]>();
          for (const r of band) lines.set(Math.round(r.v), [...(lines.get(Math.round(r.v)) ?? []), r]);
          band = [...lines.values()].sort((x, y) => y.length - x.length)[0];
        }
        if (!best || band.length > best.count) best = { page: p, count: band.length, recs: band };
      }
    }
    if (!best) return null;
    let x0 = Infinity;
    let y0 = Infinity;
    let x1 = -Infinity;
    let y1 = -Infinity;
    for (const r of best.recs) {
      const boxy = r.kind === "(" || r.kind === "h";
      x0 = Math.min(x0, r.h);
      x1 = Math.max(x1, r.h + Math.max(r.width, 2));
      y0 = Math.min(y0, boxy ? r.v - r.height : r.v - 8);
      y1 = Math.max(y1, boxy ? r.v + r.depth : r.v + 2);
    }
    return { page: best.page, x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
  }

}

/**
 * A SyncTeX input path as a path of the project: inside the run folder `work` (the main file's folder) and
 * prefixed with the main file's folder; the warm engine's wrapper and body files stand for the main file.
 */
export function projectPath(input: string, work: string, main: string, aliases: string[]): string | null {
  const norm = (p: string) => p.replace(/\\/g, "/").replace(/\/\.\//g, "/").replace(/^\.\//, "");
  const path = norm(input);
  const root = norm(work).replace(/\/$/, "");
  let rel: string | null = null;
  if (path.startsWith(root + "/")) rel = path.slice(root.length + 1);
  else if (!path.startsWith("/") && !/^[A-Za-z]:/.test(path)) rel = path;
  if (rel === null) return null;
  const name = rel.split("/").pop() ?? rel;
  const mainDir = main.includes("/") ? main.slice(0, main.lastIndexOf("/") + 1) : "";
  if (aliases.includes(name) && !rel.includes("/")) return main;
  if (!/\.[A-Za-z0-9]+$/.test(rel)) rel += ".tex";
  return mainDir + rel;
}
